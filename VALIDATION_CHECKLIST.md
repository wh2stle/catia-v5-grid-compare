# V5R27 validation checklist

Run **Debug > Compile** in the CATIA VBA editor before these checks. Use local writable files for the first pass.

Close every CATIA document before each run. The macro enforces this so only the two result products are present when CATIA tiles the windows.

## 1. Identical solids

- Create one CATPart containing a 100 × 80 × 60 mm block.
- Save a copy without changing it.
- Compare the two files.
- Expected: zero added cells and zero removed cells; both view products and all support files still save.

## 2. Added material

- Add a clearly separate 10 × 10 × 10 mm pad to the revised copy.
- Expected: blue cells only in `Added_View`; zero removed cells.
- Confirm a cell is a selectable solid Body with a Pad, not only a color overlay.

## 3. Removed material

- Remove a 10 × 10 × 10 mm region from the revised copy.
- Expected: red cells only in `Removed_View`; zero added cells.

## 4. Any-portion cell rule

- Place a change so it crosses a grid boundary by a small but unambiguous amount greater than CATIA's model tolerance.
- Expected: both intersected neighboring cells are marked.

## 5. Offset from global origin

- Position both test solids thousands of millimetres away from `(0,0,0)` while preserving their common axis system.
- Expected: the report's grid bounds surround the two actual model boxes only. They must not extend back to the global origin.

## 6. Product-instance transform

- Build two small CATProducts from the same component CATPart.
- In one revision, translate an instance in the product structure.
- Expected: added/removed cells occur at the two assembly positions. If flattening cannot paste the instance in assembly context, the macro must stop with a paste failure instead of continuing.

## 7. Bounding-box union and margin

- Read the `*_Report.txt` file.
- Verify each pre-margin minimum is the lesser input minimum and each maximum is the greater input maximum.
- Verify final span on each axis is 110% of the pre-margin span: 5% added to the negative side and 5% to the positive side.
- Verify the cell dimension is final span divided by 50.

## 8. Grid visibility and grouping

- Open either result product.
- Expected: the grid is visible initially.
- Hide `GRID_CONSTRUCTION__HIDE_THIS_COMPONENT` in the product, or `GRID_CONSTRUCTION__HIDE_THIS_GROUP` inside the grid CATPart.
- Expected: the complete lattice hides with that single action.

## 9. Appearance and views

- Expected in both products: revision 1 is opaque gray.
- Expected in Added view: actual added-cell bodies are blue with opacity 128/255.
- Expected in Removed view: actual removed-cell bodies are red with opacity 128/255.
- Rotate or zoom either tiled view before pressing Esc.
- Expected: the other view follows. Press Esc and verify they can then move independently.

## 10. Save and cleanup

- Expected: the folder dialog appears only after geometry calculation and solid-cell creation.
- Expected: seven files are created with one timestamped stem.
- Expected: no `CATIA_GridCompare_Work_*` folder from the run remains in `%TEMP%`.
- Expected: pre-existing saved comparison files are not deleted or overwritten.

## 11. 0.1 mm behavior

- Confirm a compact change whose Boolean volume is at least 0.001 mm³ is reported.
- Confirm a whole raw difference below 0.001 mm³ is rejected.
- Treat thin, broad changes as a documented exception: standard VBA cannot guarantee a local 0.1 mm thickness cutoff.

Record the CATIA V5R27 service pack/hotfix and the result of each case before production deployment.
