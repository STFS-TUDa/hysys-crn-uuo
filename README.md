# hysys-crn-uuo

VBA **User Unit Operation (UUO)** for **Aspen HYSYS** that embeds an external
chemical reactor network (CRN) gasifier model in a process flowsheet.

At each solve, the UUO reads the biomass and steam feed streams from HYSYS,
writes them to the input file of the CRN solver, runs the solver, reads the
resulting product composition and returns it to the HYSYS product stream.

This code accompanies:

> R. Di Mento et al., Development and scale-up of a kinetic reactor-network
> model for dual fluidized-bed gasification of contaminated biomass with
> process-scale integration (manuscript in preparation).

## Contents

| File | Description |
|---|---|
| `src/HYSYS_CRN_UUO.bas` | VBA source code of the User Unit Operation |
| `CITATION.cff` | Citation metadata |
| `LICENSE` | Licence of this code |

## What is *not* included

This repository contains **only the HYSYS interface**. The following are
required to run it but are **not** distributed here:

- **Aspen HYSYS** (tested with V14) and **Microsoft Excel**, on Windows
  (commercial software);
- the **CRN solver** (NetSMOKE++) and the **CRECK kinetic mechanism**, which
  must be obtained from their developers under their own licence terms;
- the CRN input files and the post-processing scripts.

## How it works

The UUO has two material feeds and one product:

| Nozzle | Name | Content |
|---|---|---|
| Feed 1 | `Bio` | solid biomass (pseudo-components) |
| Feed 2 | `Steam` | steam / gasification agent |
| Product 1 | `Prod` | product gas |

Workflow of `Execute()`:

1. read temperature, pressure, mass flow and composition of `Bio` and `Steam`;
2. update the CRN input dictionary (`inputRN.DIC`) with these values;
3. run the CRN solver through `run.bat` and wait for completion;
4. read species names and mass fractions from `Results.xlsx`
   (sheet `StreamSummary_HYSYS`);
5. write temperature, pressure, mass flow and composition to `Prod` and
   signal HYSYS that the operation is solved.

A log of each execution is written to `GasifierLink_UO.log`.

## Expected working-directory layout

The paths are defined as constants at the top of the `.bas` file and are
relative to the HYSYS case directory:

```
MapSolidNameToDic/
├── inputRN.DIC          # CRN input dictionary (updated at each solve)
├── run.bat              # launches the CRN solver
├── Results.xlsx         # solver output read back by the UUO
└── GasifierLink_UO.log  # execution log (created by the UUO)
```

The biomass pseudo-components in the HYSYS fluid package must use the same
names as the CRN mechanism (H2O in the biomass feed is mapped to `MOIST`),
in the order `CELL, XYHW, LIGC, LIGH, LIGO, TGL, TANN, MOIST, ASH, PROTC,
PROTH, PROTO`.

## Installation

1. In Aspen HYSYS, create a **User Unit Operation**.
2. Open its code editor and paste the content of `src/HYSYS_CRN_UUO.bas`.
3. Adapt the path constants if needed and connect the `Bio`, `Steam` and
   `Prod` streams and the `Heat` energy stream.

## Citation

If you use this code, please cite the article above and this repository
(see `CITATION.cff`, or the "Cite this repository" button on GitHub).

## Contact

Riccardo Di Mento — Simulation of Reactive Thermo-Fluid Systems (STFS),
Technische Universität Darmstadt — di_mento@stfs.tu-darmstadt.de
