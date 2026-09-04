# CATIA V5 R27 Grid Compare Macro

This is an importable **CATVBA source package** for CATIA V5 build 27 on 64-bit Windows. It compares either two CATParts, two CATProducts, or one of each. It does not require DMU Space Analysis or DMU Compare.

The code is split into standard VBA modules deliberately. That is useful here: geometry copying, bounding, Boolean comparison, grid scanning, output generation, saving, and camera synchronization are separate concerns, and individual modules can be tested or replaced without turning the macro into one very large script.

## What it creates

For every run, the macro creates and saves:

| File | Contents |
|---|---|
| `*_Source_Master.CATProduct` | Revision 1 and revised input, inserted one after the other |
| `*_Grid_Construction.CATPart` | Visible 50 × 50 × 50 construction lattice in `GRID_CONSTRUCTION__HIDE_THIS_GROUP` |
| `*_Added_Cells.CATPart` | Actual solid rectangular-prism bodies for added cells, pure blue and approximately 50% transparent (opacity 128/255) |
| `*_Removed_Cells.CATPart` | Actual solid rectangular-prism bodies for removed cells, pure red and approximately 50% transparent (opacity 128/255) |
| `*_Added_View.CATProduct` | Revision 1 in gray, visible grid, and blue added cells |
| `*_Removed_View.CATProduct` | Revision 1 in gray, visible grid, and red removed cells |
| `*_Report.txt` | Inputs, exact bounds, cell sizes, counts, Boolean volumes, output paths, and limitations |

The two result products are left open and tiled vertically. Their camera orientation and zoom are synchronized while the macro remains in its camera-lock loop. Press **Esc** to release the lock. The result products are saved again after the lock is released.

## Installation

CATIA cannot create a binary `.CATVBA` library from plain files outside CATIA, so this package supplies the standard importable `.bas` source modules.

1. In CATIA, open **Tools > Macro > Macros**.
2. Open **Macro libraries**, choose **VBA projects**, and create a library such as `CATIA_GridCompare.catvba`.
3. Open the Visual Basic Editor (`Alt+F11`).
4. Use **File > Import File** and import every `.bas` file from the `src` folder. Import `M_Types.bas` and `M_Config.bas` first; the rest can follow in any order.
5. In the editor, run **Debug > Compile** for the VBA project.
6. Run `M_Main.CATMain` from CATIA's macro dialog.

No additional VBA reference is required: file-system, folder-picker, and text-stream objects are late-bound. The normal CATIA V5 Automation libraries supplied with V5R27 must be available.

The CATIA installation must expose Product Structure, Part Design Boolean operations, and `HybridShapeFactory.AddNewExtremum` plus construction geometry (normally the Generative Shape Design license). **No DMU license or SPAWorkbench call is used.**

## Run behavior

1. Start from a CATIA session with no open documents. If any are open, the macro stops without changing files; this is required because CATIA's tile command acts on every document window and the result must contain exactly two panels.
2. Select revision 1 in the first CATIA file dialog.
3. Select the revised file in the second CATIA file dialog.
4. The inputs are opened in design mode as needed. If a selected input is already open with unsaved changes, the macro stops so the compared geometry and the file-backed result references cannot disagree. Product leaf instances are copied **as result without link** from the assembly selection, which bakes in their assembly positions.
5. The macro stops if a component cannot enter design mode, cannot be copied, or has a non-CATPart master shape representation. It never silently produces a partial comparison. Intentionally empty leaf products are skipped and logged.
6. Added material is computed as `revised − revision 1`; removed material is `revision 1 − revised`.
7. The union bounding box uses the actual minimum and maximum of both flattened inputs on each common axis. The global origin is not forced into the box. Five percent of that union span is added to **each side** of each axis.
8. Each axis is divided into 50 equal intervals about the union-box center. With the even count, boundary 25 is the center plane and there are 25 intervals in the negative direction and 25 in the positive direction. The resulting cells are rectangular prisms if the three axis spans differ.
9. Native Part Design Boolean intersections recursively reject empty grid regions. At leaf level, every cell containing positive difference material is marked; CATIA's Boolean kernel supplies the geometric tolerance gate.
10. If more than 25,000 solid output cells are needed for either result, the macro warns before creating that CATPart. Choosing **No** safely stops the run.
11. Only after all computation and solid-cell generation is complete does the macro ask for the output directory.
12. The result views are tiled and camera-linked. Press **Esc** to finish.

Esc can also cancel either Boolean cell scan. A canceled computation is not saved as a result.

## Default geometry settings

Edit `M_Config.bas` before compiling if a future job needs different values:

```vb
Public Const GRID_DIVISIONS_X As Long = 50
Public Const GRID_DIVISIONS_Y As Long = 50
Public Const GRID_DIVISIONS_Z As Long = 50
Public Const BOUNDS_MARGIN_FRACTION As Double = 0.05
Public Const MIN_CHANGE_MM As Double = 0.1
```

The requested defaults generate 125,000 candidate cells and a complete 3D construction lattice of 7,803 lines. CATIA file size and run time can become substantial when many cells are marked.

Keep all three division counts positive and even: an even count is what places a grid boundary exactly at the bounding-box center with the same number of intervals on each side. The macro validates these settings before opening any input.

## Important 0.1 mm limitation

V5R27's standard VBA Automation interfaces expose exact solid Booleans, but they do not expose a general, topology-independent “minimum local thickness” classifier. A true rule that rejects every change thinner than 0.1 mm in every possible direction would need a more advanced geometric implementation, normally CAA or a dedicated comparison/tolerance engine.

This macro uses the closest defensible standard-VBA rule:

- The exact added/removed Boolean result is first calculated.
- A whole result below `1E-12 m³` is ignored. That is `0.001 mm³`, the volume of a 0.1 mm cube.
- Once that result passes, the any-portion rule is honored per cell: any positive Boolean-intersection volume marks the cell.

This suppresses minute Boolean noise, but it is **not equivalent to a guaranteed 0.1 mm thickness filter**. A broad sheet-like change thinner than 0.1 mm can exceed the volume threshold and be reported. The run report repeats this limitation.

## Input and environment limits

- The macro requires the Part Design and Generative Shape Design automation capabilities described above; it does not require DMU Space Analysis or DMU Compare.
- Inputs must contain solid Part Design geometry. Surface-only, unloaded, broken-link, or non-CATPart leaf representations cause the run to stop; intentionally empty product leaves are skipped and logged.
- Both revisions must already use the same model or assembly axis system. The macro does not register, align, or best-fit the revisions.
- Every referenced CATPart in a CATProduct must be reachable and loadable in design mode.
- CATIA's built-in `Windows.Arrange` method affects every document window, so the macro requires an empty document session at startup and leaves only the two result products open for tiling.
- Camera synchronization is live VBA behavior, not a permanent associativity stored between two CATProduct files. It runs until Esc is pressed.
- The macro removes only its own `%TEMP%\CATIA_GridCompare_Work_*` folders and closes its temporary documents. It deliberately does not purge CATIA's shared cache or touch prior saved result files.

## Recommended first validation

Before using production assemblies, run the cases in `docs/VALIDATION_CHECKLIST.md`. Begin with two simple blocks, then a small assembly with one translated component. This confirms the local V5R27 hotfix level, licenses, CATSettings, and Paste Special behavior.

## Source modules

| Module | Responsibility |
|---|---|
| `M_Main` | End-to-end orchestration and safe cleanup |
| `M_Types`, `M_Config` | Shared data structures and defaults |
| `M_IO_Products` | File dialogs, source master, result products, output directory |
| `M_CopyFlatten` | CATPart/CATProduct flattening in assembly position |
| `M_Bounds` | Exact axis extrema, union, margin, and cell sizes |
| `M_BooleanDifference` | Added and removed Part Design solids |
| `M_GridScan` | Adaptive Boolean cell intersection scan |
| `M_SolidGeometry` | Offset planes, sketches, pads, and probe prisms |
| `M_OutputGeometry` | Full construction lattice and colored solid cell CATParts |
| `M_ViewsCamera` | Vertical tiling and live viewpoint synchronization |
| `M_Report`, `M_Utilities` | Audit report, appearance, measurement, temp cleanup, and helpers |
