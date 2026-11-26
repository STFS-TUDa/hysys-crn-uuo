Option Explicit

' =========================================================
' FILE PATHS AND CONFIGURATION CONSTANTS
' All paths point to the CRN gasifier working directory.
' =========================================================

' Path to the CRN input dictionary file (updated at each solve)
Const DIC_FILE As String = "MapSolidNameToDic\inputRN.DIC"

' Path to the batch file that launches the CRN simulation
Const RUN_BAT  As String = "MapSolidNameToDic\run.bat"

' Path to the Excel file written by the CRN solver with output stream composition
Const RESULTS_XLSX  As String = "MapSolidNameToDic\Results.xlsx"

' Name of the worksheet inside the Excel file containing the results
Const RESULTS_SHEET As String = "StreamSummary_HYSYS"

' Row/column layout of the results table in the Excel worksheet:
'   - Data starts at row 46
'   - Column N (14) contains species names
'   - Column O (15) contains mass fraction values
Const RESULTS_START_ROW As Long = 46
Const RESULTS_MAX_ROWS  As Long = 200
Const RESULTS_COL_NAME  As Long = 14  ' Column N: species name
Const RESULTS_COL_VAL   As Long = 15  ' Column O: mass fraction value

' Optional: if True, writes a simplification cut-off value to a cell
' in the Excel file before reading results (used to filter trace species)
Const USE_SIMPLIFICATION_CUT As Boolean = False
Const RESULTS_CELL_CUT As String = "B42"
Const SIMPLIFICATION_CUT As Double = 0.0000001

' Path to the log file written during execution for debugging
Const LOG_FILE As String = "MapSolidNameToDic\GasifierLink_UO.log"

' =========================================================
' SPECIES NAME MAPPING NOTES (Stream 1 - Solid Biomass)
'
' HYSYS name  -->  CRN DIC name
'   H2O       -->  MOIST
'   All others keep the same name
'
' Required order in the DIC file:
'   CELL, XYHW, LIGC, LIGH, LIGO, TGL, TANN, MOIST, ASH, PROTC, PROTH, PROTO
' =========================================================

' Guard flag to prevent re-entrant calls to Execute()
' (HYSYS may call Execute multiple times during a single solve pass)
Dim g_isExecuting As Boolean

' =========================================================
' Initialize
' Called once when the unit operation is loaded or reset.
' Configures the nozzle names and activates the correct ports.
' =========================================================
Sub Initialize()
On Error GoTo ErrorTrap

    ' Define the two material feed nozzles: biomass (Feed1) and steam (Feed2)
    ActiveObject.Feeds1Name = "Bio"
    ActiveObject.Feeds2Name = "Steam"
    ActiveObject.Feeds2Active = True         ' Feed2 must be explicitly activated

    ' Define the single product nozzle
    ActiveObject.Products1Name = "Prod"

    ' Activate the energy feed nozzle for heat duty calculation
    ' Heat duty is computed manually from enthalpy balance in Execute()
    ActiveObject.EnergyFeedsActive = True
    ActiveObject.EnergyFeedsName = "Heat"

    ' No energy product nozzle needed
    ActiveObject.EnergyProductsActive = False

    ' Reset re-entrancy guard
    g_isExecuting = False
    LogMsg "Initialize OK"

    Exit Sub
ErrorTrap:
    MsgBox "Init Error: " & Err.Description
End Sub

' =========================================================
' Execute
' Called by the HYSYS solver when all required inputs are
' available and the unit operation needs to be calculated.
'
' Workflow:
'   1. Read T, P, MassFlow and composition from Bio and Steam feeds
'   2. Update the CRN DIC input file with current stream data
'   3. Run the CRN solver via the batch file and wait for completion
'   4. Read output composition from the Excel results file
'   5. Write T, P, MassFlow and composition to the Product stream
'   6. Signal solver that the unit operation is fully solved
' =========================================================
Sub Execute()
On Error GoTo ErrorTrap

    ' Skip the forgetting pass (HYSYS calls Execute twice: once to clear, once to calculate)
    If ActiveObject.SimulationCase.Solver.IsForgetting = True Then Exit Sub

    ' Skip if the unit operation is set to Ignored
    If ActiveObject.IsIgnored = True Then Exit Sub

    ' Prevent re-entrant execution
    If g_isExecuting = True Then
        LogMsg "Execute skipped (re-entrancy)"
        Exit Sub
    End If
    g_isExecuting = True

    ' --- Get stream objects from the three nozzles ---
    Dim bio As Object, steam As Object, prod As Object
    Set bio   = GetFirstStream(ActiveObject.Feeds1)    ' Biomass feed
    Set steam = GetFirstStream(ActiveObject.Feeds2)    ' Steam feed
    Set prod  = GetFirstStream(ActiveObject.Products1) ' Product gas

    ' Exit early if any stream is not connected
    If bio   Is Nothing Then GoTo EarlyExit
    If steam Is Nothing Then GoTo EarlyExit
    If prod  Is Nothing Then GoTo EarlyExit

    ' Exit early if Bio stream conditions are not yet specified
    If Not bio.Temperature.IsKnown Then GoTo EarlyExit
    If Not bio.Pressure.IsKnown    Then GoTo EarlyExit
    If Not bio.MassFlow.IsKnown    Then GoTo EarlyExit

    ' Exit early if Steam stream conditions are not yet specified
    If Not steam.Temperature.IsKnown Then GoTo EarlyExit
    If Not steam.Pressure.IsKnown    Then GoTo EarlyExit
    If Not steam.MassFlow.IsKnown    Then GoTo EarlyExit

    ' --- Read operating conditions from the Steam stream ---
    ' Temperature and pressure are taken from Steam (reactor operating conditions)
    Dim T As Double, PkPa As Double
    T    = steam.Temperature.GetValue("K")
    PkPa = steam.Pressure.GetValue("kPa")

    ' --- Read mass flow rates ---
    Dim MF_bio As Double, MF_steam As Double
    MF_bio   = bio.MassFlow.GetValue("kg/s")
    MF_steam = steam.MassFlow.GetValue("kg/s")

    ' --- Read mass fraction arrays (indexed by fluid package component order) ---
    Dim zBio, zSteam
    zBio   = bio.ComponentMassFractionValue
    zSteam = steam.ComponentMassFractionValue

    LogMsg "Execute start: T=" & T & " K, P=" & PkPa & " kPa, MFbio=" & MF_bio & " kg/s, MFsteam=" & MF_steam & " kg/s"

    ' --- Step 1: Update the CRN DIC input file ---
    ' Writes SystemPressure, Temperature, MassFlowRate and compositions
    ' for both Stream 1 (solid biomass) and Stream 2 (steam/gas)
    Call UpdateDicFile(T, PkPa, MF_bio, bio, zBio, MF_steam, steam, zSteam)

    ' --- Step 2: Run the CRN simulation ---
    ' Executes run.bat synchronously (waits for completion)
    Call RunBat(RUN_BAT)

    ' --- Step 3: Read output composition from Excel ---
    ' Reads species names and mass fractions from the StreamSummary_HYSYS sheet
    Dim resNames, resVals
    Call ReadResultsFromExcel(RESULTS_XLSX, RESULTS_SHEET, resNames, resVals)

    If IsEmpty(resNames) Then
        LogMsg "No Excel results read -> exit"
        GoTo EarlyExit
    End If

    ' --- Step 4: Write results to the Product stream ---
    ' T and P are inherited from Steam (reactor outlet assumed at same conditions)
    ' MassFlow = Bio + Steam (mass conservation)
    prod.Pressure.Calculate    steam.PressureValue
    prod.Temperature.Calculate steam.TemperatureValue
    prod.MassFlow.Calculate    (MF_bio + MF_steam)

    ' Apply species composition by matching names from Excel to HYSYS component list
    Call ApplyCompositionByName_MassFraction(prod, resNames, resVals)

    ' --- Step 5: Signal HYSYS that this unit operation is fully solved ---
    ActiveObject.SolveComplete
    LogMsg "Execute OK"

EarlyExit:
    g_isExecuting = False
    Exit Sub

ErrorTrap:
    LogMsg "Execute ERROR: " & Err.Number & " - " & Err.Description
    g_isExecuting = False
    MsgBox "Execute Error: " & Err.Description
End Sub

' =========================================================
' StatusQuery
' Called by HYSYS to determine the status indicator colour.
' Adds warning messages if required streams are not connected.
' =========================================================
Sub StatusQuery()
On Error GoTo ErrorTrap

    If ActiveObject.IsIgnored = True Then Exit Sub

    Dim bio As Object, steam As Object, prod As Object
    Set bio   = GetFirstStream(ActiveObject.Feeds1)
    Set steam = GetFirstStream(ActiveObject.Feeds2)
    Set prod  = GetFirstStream(ActiveObject.Products1)

    If bio   Is Nothing Then Call ActiveObject.AddStatusCondition(slMissingRequiredInformation, 1, "Requires Bio feed (Feed1)")
    If steam Is Nothing Then Call ActiveObject.AddStatusCondition(slMissingRequiredInformation, 2, "Requires Steam feed (Feed2)")
    If prod  Is Nothing Then Call ActiveObject.AddStatusCondition(slMissingRequiredInformation, 3, "Requires Product stream (Product1)")

    Exit Sub
ErrorTrap:
    MsgBox "Status Query Error: " & Err.Description
End Sub

' =========================================================
' UpdateDicFile
' Reads the DIC file line by line and updates:
'   - SystemPressure (converted from kPa to Pa)
'   - All Temperature lines (in K)
'   - @Stream 1 block (solid biomass: flow + composition)
'   - @Stream 2 block (steam/gas: flow + composition)
' Then writes the modified content back to disk.
' =========================================================
Sub UpdateDicFile(T As Double, PkPa As Double, MF_bio As Double, bio As Object, zBio, MF_steam As Double, steam As Object, zSteam)
On Error GoTo ErrorTrap

    If Not FileExists(DIC_FILE) Then
        LogMsg "DIC not found: " & DIC_FILE
        Exit Sub
    End If

    Dim lines
    lines = ReadAllLines(DIC_FILE)

    ' Update SystemPressure: convert kPa -> Pa
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If StartsWith(Trim(lines(i)), "SystemPressure") Then
            lines(i) = "SystemPressure " & FormatNumberInvariant(PkPa * 1000#, 2) & " Pa"
        End If
    Next i

    ' Update all Temperature entries (reactor temperature, in Kelvin)
    For i = LBound(lines) To UBound(lines)
        If StartsWith(Trim(lines(i)), "Temperature") Then
            lines(i) = "Temperature " & FormatNumberInvariant(T, 2) & " K"
        End If
    Next i

    ' Update the solid biomass stream block (@Stream 1)
    Call UpdateStream1Block(lines, MF_bio, bio, zBio)

    ' Update the steam/gas stream block (@Stream 2)
    Call UpdateStream2Block(lines, MF_steam, steam, zSteam)

    Call WriteAllLines(DIC_FILE, lines)
    LogMsg "DIC updated successfully"

    Exit Sub
ErrorTrap:
    LogMsg "UpdateDicFile ERROR: " & Err.Description
End Sub

' =========================================================
' UpdateStream1Block
' Updates the @Stream 1 section in the DIC file:
'   - MassFlowRateSolid: total biomass mass flow [kg/s]
'   - MassFractionSolid: per-species solid composition
'     (species names mapped from HYSYS to DIC nomenclature,
'      written in the fixed order required by the CRN solver)
' Component names are read from the HYSYS fluid package,
' not from the stream object (which does not expose them directly).
' =========================================================
Sub UpdateStream1Block(ByRef lines, MF_bio As Double, bio As Object, zBio)
On Error GoTo ErrorTrap

    ' Locate the @Stream 1 block boundaries
    Dim startIdx As Long, endIdx As Long
    If Not FindStreamBlock(lines, 1, startIdx, endIdx) Then
        LogMsg "WARNING: @Stream 1 block not found"
        Exit Sub
    End If

    ' Update MassFlowRateSolid line
    Dim i As Long
    For i = startIdx To endIdx
        If StartsWith(Trim(lines(i)), "MassFlowRateSolid") Then
            lines(i) = "MassFlowRateSolid " & FormatNumberInvariant(MF_bio, 14) & " kg/s"
        End If
    Next i

    ' Find the start of the MassFractionSolid block
    Dim mfStart As Long, mfEnd As Long
    mfStart = -1: mfEnd = -1

    For i = startIdx To endIdx
        If StartsWith(Trim(lines(i)), "MassFractionSolid") Then
            mfStart = i
            Exit For
        End If
    Next i

    If mfStart < 0 Then
        LogMsg "WARNING: MassFractionSolid not found in @Stream 1"
        Exit Sub
    End If

    ' Find the end of the MassFractionSolid block (terminated by ';')
    For i = mfStart To endIdx
        If InStr(CStr(lines(i)), ";") > 0 Then
            mfEnd = i
            Exit For
        End If
    Next i

    If mfEnd < 0 Then
        LogMsg "WARNING: MassFractionSolid terminator ';' not found in @Stream 1"
        Exit Sub
    End If

    ' Build component name array from the HYSYS fluid package
    ' (stream object does not expose ComponentListNames directly)
    Dim compNames
    Dim comps As Object
    Set comps = ActiveObject.Flowsheet.FluidPackage.Components
    Dim nC As Long
    nC = comps.Count
    ReDim compNames(0 To nC - 1)
    Dim ic As Long
    For ic = 0 To nC - 1
        compNames(ic) = comps.Item(ic).name
    Next ic

    ' Build the new MassFractionSolid lines (mapped + ordered for the CRN solver)
    Dim newBlock
    newBlock = BuildSolidMassFractionLines(compNames, zBio)

    ' Replace the old MassFractionSolid block with the new one
    Dim rebuilt()
    rebuilt = ReplaceLinesRange(lines, mfStart, mfEnd, newBlock)
    lines = rebuilt

    Exit Sub
ErrorTrap:
    LogMsg "UpdateStream1Block ERROR: " & Err.Description
End Sub

' =========================================================
' UpdateStream2Block
' Updates the @Stream 2 section in the DIC file:
'   - MassFlowRate: steam mass flow [kg/s]
'   - MassFraction: per-species gas composition
'     (component names from HYSYS fluid package)
' =========================================================
Sub UpdateStream2Block(ByRef lines, MF_steam As Double, steam As Object, zSteam)
On Error GoTo ErrorTrap

    ' Locate the @Stream 2 block boundaries
    Dim startIdx As Long, endIdx As Long
    If Not FindStreamBlock(lines, 2, startIdx, endIdx) Then
        LogMsg "WARNING: @Stream 2 block not found"
        Exit Sub
    End If

    ' Update MassFlowRate line
    Dim i As Long
    For i = startIdx To endIdx
        If StartsWith(Trim(lines(i)), "MassFlowRate") Then
            lines(i) = "MassFlowRate " & FormatNumberInvariant(MF_steam, 6) & " kg/s"
        End If
    Next i

    ' Find the start of the MassFraction block
    Dim mfStart As Long, mfEnd As Long
    mfStart = -1: mfEnd = -1

    For i = startIdx To endIdx
        If StartsWith(Trim(lines(i)), "MassFraction") Then
            mfStart = i
            Exit For
        End If
    Next i

    If mfStart < 0 Then
        LogMsg "WARNING: MassFraction not found in @Stream 2"
        Exit Sub
    End If

    ' Find the end of the MassFraction block (terminated by ';')
    For i = mfStart To endIdx
        If InStr(1, lines(i), ";") > 0 Then
            mfEnd = i
            Exit For
        End If
    Next i

    If mfEnd < 0 Then
        LogMsg "WARNING: MassFraction terminator ';' not found in @Stream 2"
        Exit Sub
    End If

    ' Build component name array from the HYSYS fluid package
    Dim compNames
    Dim comps As Object
    Set comps = ActiveObject.Flowsheet.FluidPackage.Components
    Dim nC As Long
    nC = comps.Count
    ReDim compNames(0 To nC - 1)
    Dim ic As Long
    For ic = 0 To nC - 1
        compNames(ic) = comps.Item(ic).name
    Next ic

    ' Build the new MassFraction lines
    Dim newBlock
    newBlock = BuildGasMassFractionLines(compNames, zSteam)

    ' Replace the old MassFraction block with the new one
    Dim rebuilt()
    rebuilt = ReplaceLinesRange(lines, mfStart, mfEnd, newBlock)
    lines = rebuilt

    Exit Sub
ErrorTrap:
    LogMsg "UpdateStream2Block ERROR: " & Err.Description
End Sub

' =========================================================
' FindStreamBlock
' Scans the DIC lines array to find the start and end indices
' of a "@Stream N" block. The block ends just before the next
' "@Stream" tag or at the end of the file.
' Returns True if found, False otherwise.
' =========================================================
Function FindStreamBlock(lines, streamNumber As Long, ByRef startIdx As Long, ByRef endIdx As Long) As Boolean
    Dim i As Long
    Dim tag As String
    tag = "@Stream " & CStr(streamNumber)

    startIdx = -1: endIdx = -1

    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) = tag Then
            startIdx = i
            Exit For
        End If
    Next i

    If startIdx < 0 Then
        FindStreamBlock = False
        Exit Function
    End If

    ' Default: block extends to end of file
    endIdx = UBound(lines)
    ' Trim at next @Stream tag if present
    For i = startIdx + 1 To UBound(lines)
        If StartsWith(Trim(lines(i)), "@Stream ") Then
            endIdx = i - 1
            Exit For
        End If
    Next i

    FindStreamBlock = True
End Function

' =========================================================
' MapSolidNameToDic
' Maps HYSYS component names to CRN DIC species names
' for the solid biomass stream (Stream 1).
' All species keep the same name, except H2O, which is
' renamed to MOIST as required by the DIC format.
' Returns empty string for components not present in Stream 1.
' =========================================================
Function MapSolidNameToDic(cleanUpper As String) As String
    Select Case cleanUpper
        Case "CELL": MapSolidNameToDic = "CELL"
        Case "XYHW": MapSolidNameToDic = "XYHW"
        Case "LIGC": MapSolidNameToDic = "LIGC"
        Case "LIGO": MapSolidNameToDic = "LIGO"  ' no swap needed (with respect to the phyton version)
        Case "LIGH": MapSolidNameToDic = "LIGH"  ' no swap needed
        Case "TGL":  MapSolidNameToDic = "TGL"
        Case "TANN": MapSolidNameToDic = "TANN"
        Case "H2O":  MapSolidNameToDic = "MOIST" ' renamed
        Case "ASH":  MapSolidNameToDic = "ASH"
        Case "PROTC": MapSolidNameToDic = "PROTC"
        Case "PROTH": MapSolidNameToDic = "PROTH"
        Case "PROTO": MapSolidNameToDic = "PROTO"
        Case Else:   MapSolidNameToDic = ""       ' not a solid species -> skip
    End Select
End Function

' =========================================================
' BuildSolidMassFractionLines
' Builds the MassFractionSolid block for the DIC file.
' - Maps HYSYS component names to DIC names via MapSolidNameToDic
' - Skips components with negligible mass fraction (< 1e-10)
' - Writes species in the fixed order required by the CRN solver
' - Appends ';' to the last line as DIC block terminator
' - Falls back to ASH=1 if no valid species are found
' =========================================================
Function BuildSolidMassFractionLines(compNames, zBio) As Variant
    ' Use a dictionary for fast lookup by DIC name
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1 ' case-insensitive

    Dim i As Long
    For i = LBound(compNames) To UBound(compNames)
        Dim nm As String
        nm = CStr(compNames(i))
        nm = Replace(nm, "*", "")  ' remove HYSYS solid phase marker
        nm = UCase(Trim(nm))

        Dim dicName As String
        dicName = MapSolidNameToDic(nm)

        If dicName <> "" Then
            Dim v As Double
            v = CDbl(zBio(i))
            If v > 1E-10 Then dic(dicName) = v  ' skip trace/zero fractions
        End If
    Next i

    ' Fixed species order as required by the CRN DIC format
    Dim orderArr
    orderArr = Array("CELL", "XYHW", "LIGC", "LIGH", "LIGO", "TGL", "TANN", "MOIST", "ASH", "PROTC", "PROTH", "PROTO")

    Dim outLines() As String
    Dim k As Long, n As Long
    n = -1

    For k = 0 To UBound(orderArr)
        Dim key As String
        key = CStr(orderArr(k))
        If dic.Exists(key) Then
            n = n + 1
            ReDim Preserve outLines(0 To n)
            If n = 0 Then
                ' First line includes the keyword
                outLines(n) = "MassFractionSolid     " & key & "  " & Format(CDbl(dic(key)), "0.00000000000000")
            Else
                ' Continuation lines are indented to align with the first species value
                outLines(n) = "                      " & key & "  " & Format(CDbl(dic(key)), "0.00000000000000")
            End If
        End If
    Next k

    ' Append DIC block terminator ';' to the last line
    If n >= 0 Then
        outLines(n) = outLines(n) & ";"
        BuildSolidMassFractionLines = outLines
    Else
        ' Fallback: write a minimal valid block to avoid corrupting the DIC file
        ReDim outLines(0 To 0)
        outLines(0) = "MassFractionSolid     ASH  1.00000000000000;"
        BuildSolidMassFractionLines = outLines
    End If
End Function

' =========================================================
' BuildGasMassFractionLines
' Builds the MassFraction block for the gas stream (Stream 2).
' - Uses component names directly (no name mapping needed for gas)
' - Skips components with negligible mass fraction (< 1e-10)
' - Appends ';' to the last line as DIC block terminator
' - Falls back to H2O=1 if no valid species are found
' =========================================================
Function BuildGasMassFractionLines(compNames, zSteam) As Variant
    Dim outLines() As String
    Dim n As Long
    n = -1

    Dim i As Long
    For i = LBound(compNames) To UBound(compNames)
        Dim v As Double
        v = CDbl(zSteam(i))
        If v > 1E-10 Then
            Dim nm As String
            nm = CStr(compNames(i))
            nm = Replace(nm, "*", "")  ' remove HYSYS solid phase marker if present
            nm = Trim(nm)

            n = n + 1
            ReDim Preserve outLines(0 To n)
            If n = 0 Then
                outLines(n) = "MassFraction      " & nm & "  " & Format(v, "0.00000000000000")
            Else
                outLines(n) = "                      " & nm & "  " & Format(v, "0.00000000000000")
            End If
        End If
    Next i

    If n >= 0 Then
        outLines(n) = outLines(n) & ";"
        BuildGasMassFractionLines = outLines
    Else
        ' Fallback: pure steam
        ReDim outLines(0 To 0)
        outLines(0) = "MassFraction      H2O  1.00000000000000;"
        BuildGasMassFractionLines = outLines
    End If
End Function

' =========================================================
' RunBat
' Executes the CRN batch file synchronously.
' Uses cmd /c to ensure .bat files are correctly interpreted.
' Sets the working directory to the folder containing the .bat
' so that all relative paths inside run.bat resolve correctly.
' The shell window is hidden (style=0) and execution waits
' for the process to finish (bWait=True) before returning.
' =========================================================
Sub RunBat(batPath As String)
On Error GoTo ErrorTrap

    If Not FileExists(batPath) Then
        LogMsg "BAT not found: " & batPath
        Exit Sub
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    ' Set working directory to the folder containing the .bat file
    ' This ensures the CRN solver executable and all relative paths resolve correctly
    Dim batFolder As String
    batFolder = Left(batPath, InStrRev(batPath, "\"))
    sh.CurrentDirectory = batFolder

    LogMsg "Running: " & batPath & " from folder: " & batFolder
    Dim rc As Long
    ' cmd /c is required to correctly launch .bat files via WScript.Shell
    rc = sh.Run("cmd /c """ & batPath & """", 0, True) ' 0 = hidden window, True = wait for exit
    LogMsg "run.bat exit code=" & rc

    Exit Sub
ErrorTrap:
    LogMsg "RunBat ERROR: " & Err.Description
End Sub
' =========================================================
' ReadResultsFromExcel
' Opens the CRN results Excel file (read-only) and reads
' species names and mass fraction values from the results table.
' Returns two arrays: names() and vals() of equal length.
' Reading stops at the first empty cell in the name column.
' Excel is closed and quit after reading to free resources.
' =========================================================
Sub ReadResultsFromExcel(xlsxPath As String, sheetName As String, ByRef names, ByRef vals)
On Error GoTo ErrorTrap

    names = Empty
    vals  = Empty

    If Not FileExists(xlsxPath) Then
        LogMsg "Results.xlsx not found: " & xlsxPath
        Exit Sub
    End If

    ' Open Excel invisibly
    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    xl.Visible = False

    Set wb = xl.Workbooks.Open(xlsxPath, False, True) ' read-only
    Set ws = wb.Worksheets(sheetName)

    ' Optionally write a simplification cut-off value before reading
    ' (used to zero out species below a threshold in the Excel model)
    If USE_SIMPLIFICATION_CUT = True Then
        ws.Range(RESULTS_CELL_CUT).Value = SIMPLIFICATION_CUT
        wb.Save
    End If

    Dim tmpN() As String
    Dim tmpV() As Double
    Dim k As Long
    k = -1

    ' Read rows until an empty name cell is found
    Dim r As Long
    For r = RESULTS_START_ROW To (RESULTS_START_ROW + RESULTS_MAX_ROWS - 1)
        Dim nm As String
        nm = Trim(CStr(ws.Cells(r, RESULTS_COL_NAME).Value))
        If Len(nm) = 0 Then Exit For ' end of data

        Dim vv As Variant
        vv = ws.Cells(r, RESULTS_COL_VAL).Value

        If IsNumeric(vv) Then
            k = k + 1
            ReDim Preserve tmpN(0 To k)
            ReDim Preserve tmpV(0 To k)
            tmpN(k) = nm
            tmpV(k) = CDbl(vv)
        Else
            LogMsg "Excel row " & r & ": non numeric value for " & nm & " -> skipped"
        End If
    Next r

    ' Close Excel cleanly
    wb.Close False
    xl.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing

    If k >= 0 Then
        names = tmpN
        vals  = tmpV
        LogMsg "Excel read OK. n=" & (k + 1)
    Else
        LogMsg "Excel read returned 0 species."
    End If

    Exit Sub
ErrorTrap:
    LogMsg "ReadResultsFromExcel ERROR: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Sub

' =========================================================
' ApplyCompositionByName_MassFraction
' Writes mass fraction composition to a HYSYS stream object.
' - Matches species names from the results array to HYSYS
'   component indices using ComponentIndexByName()
' - Unmatched species are silently ignored
' - Normalises the fraction array before writing to ensure
'   the sum equals 1.0 (handles rounding from the CRN solver output)
' =========================================================
Sub ApplyCompositionByName_MassFraction(streamObj As Object, names, vals)
On Error GoTo ErrorTrap

    ' Get the component list from the active fluid package
    Dim comps As Object
    Set comps = ActiveObject.Flowsheet.FluidPackage.Components

    Dim nComp As Long
    nComp = comps.Count

    ' Initialise fraction array to zero
    Dim fracs()
    ReDim fracs(0 To nComp - 1)

    ' Map each result species to its HYSYS component index
    Dim i As Long
    For i = 0 To UBound(names)
        Dim sName As String
        sName = UCase(Trim(CStr(names(i))))

        Dim idx As Long
        idx = ComponentIndexByName(comps, sName)
        If idx >= 0 Then
            fracs(idx) = CDbl(vals(i))
        End If
        ' Species not found in HYSYS component list are ignored
    Next i

    ' Normalise to ensure fractions sum to exactly 1.0
    Dim s As Double
    s = 0#
    Dim j As Long
    For j = 0 To nComp - 1
        s = s + fracs(j)
    Next j
    If s > 0# Then
        For j = 0 To nComp - 1
            fracs(j) = fracs(j) / s
        Next j
    End If

    ' Write the normalised mass fraction array to the stream
    streamObj.ComponentMassFraction.Calculate fracs

    Exit Sub
ErrorTrap:
    LogMsg "ApplyCompositionByName ERROR: " & Err.Description
End Sub

' =========================================================
' ComponentIndexByName
' Returns the zero-based index of a component in the HYSYS
' fluid package component list, matched by uppercase name.
' First tries the native comps.index() method for speed;
' falls back to a linear search if that fails.
' Returns -1 if the component is not found.
' =========================================================
Function ComponentIndexByName(comps As Object, upperName As String) As Long
    On Error Resume Next

    ' Fast path: use the built-in index lookup
    Dim idx As Variant
    idx = comps.index(upperName)
    If Err.Number = 0 Then
        ComponentIndexByName = CLng(idx)
        Exit Function
    End If
    Err.Clear

    ' Fallback: linear search by name
    Dim i As Long
    For i = 0 To comps.Count - 1
        If UCase(CStr(comps.Item(i).name)) = upperName Then
            ComponentIndexByName = i
            Exit Function
        End If
    Next i

    ComponentIndexByName = -1 ' not found
End Function

' =========================================================
' GetFirstStream
' Returns the first stream object from a HYSYS nozzle
' collection (Feeds1, Feeds2, Products1, etc.).
' Returns Nothing if the collection is empty (not connected).
' =========================================================
Function GetFirstStream(col As Object) As Object
On Error Resume Next
    If col.Count > 0 Then
        Set GetFirstStream = col.Item(0)
    Else
        Set GetFirstStream = Nothing
    End If
End Function

' =========================================================
' ReplaceLinesRange
' Returns a new array with lines[fromIdx..toIdx] replaced
' by the contents of newLines. Used to splice updated
' composition blocks into the DIC file line array.
' =========================================================
Function ReplaceLinesRange(lines, fromIdx As Long, toIdx As Long, newLines) As Variant
    Dim out()
    Dim nOut As Long
    nOut = -1

    Dim i As Long

    ' Copy lines before the replaced range
    For i = LBound(lines) To fromIdx - 1
        nOut = nOut + 1
        ReDim Preserve out(0 To nOut)
        out(nOut) = lines(i)
    Next i

    ' Insert the new lines
    Dim k As Long
    For k = LBound(newLines) To UBound(newLines)
        nOut = nOut + 1
        ReDim Preserve out(0 To nOut)
        out(nOut) = newLines(k)
    Next k

    ' Copy lines after the replaced range
    For i = toIdx + 1 To UBound(lines)
        nOut = nOut + 1
        ReDim Preserve out(0 To nOut)
        out(nOut) = lines(i)
    Next i

    ReplaceLinesRange = out
End Function

' =========================================================
' FILE I/O UTILITIES
' =========================================================

' Returns True if the file at path p exists
Function FileExists(p As String) As Boolean
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function

' Reads all lines of a text file into a string array (split on CRLF)
Function ReadAllLines(p As String) As Variant
    Dim txt As String
    txt = ReadAllText(p)
    ReadAllLines = Split(txt, vbCrLf)
End Function

' Joins a string array with CRLF and writes to file
Sub WriteAllLines(p As String, lines)
    Dim txt As String
    txt = Join(lines, vbCrLf)
    WriteAllText p, txt
End Sub

' Reads the entire content of a text file as a single string
Function ReadAllText(p As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(p, 1, False)
    ReadAllText = ts.ReadAll
    ts.Close
End Function

' Writes a string to a file (overwrites existing content)
Sub WriteAllText(p As String, txt As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(p, 2, True)
    ts.Write txt
    ts.Close
End Sub

' =========================================================
' STRING / NUMBER UTILITIES
' =========================================================

' Returns True if string s starts with the given prefix
Function StartsWith(s, prefix)
    StartsWith = (Left(CStr(s), Len(CStr(prefix))) = CStr(prefix))
End Function

' Formats a number with nDec decimal places using invariant (dot) separator
' Needed because VBA uses the locale decimal separator (comma in Italian locale)
Function FormatNumberInvariant(x, nDec)
    Dim s
    s = CStr(Round(CDbl(x), CLng(nDec)))
    FormatNumberInvariant = Replace(s, ",", ".")
End Function

' =========================================================
' LogMsg
' Appends a timestamped message to the log file.
' Uses append mode (8) so previous entries are preserved.
' Errors are silently ignored to avoid masking the main logic.
' =========================================================
Sub LogMsg(msg As String)
On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(LOG_FILE, 8, True) ' 8 = append
    ts.WriteLine "[" & Now & "] " & msg
    ts.Close
End Sub
