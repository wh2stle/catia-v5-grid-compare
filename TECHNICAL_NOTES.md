# Technical notes

## Comparison definition

- Added material: `revised − revision 1`
- Removed material: `revision 1 − revised`

The macro uses Part Design `Add`, `Assemble` fallback, `Remove`, and temporary `Intersect` operations through `ShapeFactory`. It does not start interactive CATIA commands and does not automate DMU Compare.

Volume is read from the new working CATPart's main PartBody with Product Structure `Product.Analyze.Volume`. Bounding coordinates come from explicit datums of GSD extremum points. Consequently, the implementation never requests `SPAWorkbench`; it does require the Part Design and Generative Shape Design automation capabilities.

## Assembly flattening

Every terminal CATPart instance in a CATProduct is switched to design mode and selected in the source ProductDocument. It is copied and pasted into a temporary CATPart using `CATPrtResultWithOutLink`. Selecting the assembly instance rather than its reference Part is intentional: the instance transform is captured in the common assembly coordinate system.

If design-mode loading or any CATPart instance paste fails, the run aborts. A non-CATPart leaf with a master shape representation also aborts; only an intentionally empty product leaf is skipped. This protects against plausible-looking but incomplete output when a referenced component is unloaded, inaccessible, or unsupported.

## Bounding box

Six `HybridShapeExtremum` objects per flattened result body find minimum and maximum X, Y, and Z values. Each uses the requested axis as its primary direction and the other two axes as tie-break directions, because a one-direction extremum of a planar solid can be a face rather than a single point. A zero-offset coordinate point referenced to the resulting three-direction extremum exposes its Cartesian coordinates without SPAWorkbench. The temporary extremum set is deleted immediately after measurement. The algorithm unions both revisions by taking the lesser minimum and greater maximum on each axis; it never seeds the result with zero.

For an unpadded span `L`, final bounds are `min − 0.05L` and `max + 0.05L`. Therefore the final span is `1.10L`, centered on the union box center.

## Cell search

Testing 125,000 cells naively would require 125,000 expensive Part Design operations. The scanner starts with the complete index range and recursively bisects the physically longest remaining range:

- Empty intersection: discard the whole range.
- Measured intersection is at least the complete mathematical range volume: mark every contained leaf cell without more probes.
- Partial intersection: split the range and continue.
- One cell: mark it.

Each probe is a true padded rectangular-prism Body intersected with the exact difference Body. The temporary Boolean feature, Body, plane, sketch, and pad are then deleted. At intervals, the difference volume is remeasured to verify that deletion restored the source Body.

The scanner deliberately has no "nearly full" relative shortcut. If numerical rounding puts a full range just below its mathematical volume, the macro subdivides further; this costs time but avoids painting a genuinely empty cell inside a small cavity.

## Output geometry

Each marked cell becomes its own named Body and Pad in the added or removed CATPart. This satisfies the requirement for actual selectable solids. Fifty reusable offset-plane supports avoid generating a separate plane for every cell.

The construction file contains every lattice line, not only an outer frame:

- 2,601 X-parallel lines (`51 × 51`)
- 2,601 Y-parallel lines (`51 × 51`)
- 2,601 Z-parallel lines (`51 × 51`)
- 7,803 lines total

All lattice lines are under `GRID_CONSTRUCTION__HIDE_THIS_GROUP`. Support points are in a hidden nested set.

## Automation references consulted

- CATIA V5 Automation object documentation: <https://catiadesign.org/_doc/V5Automation/>
- ShapeFactory methods derived from CAA V5 Visual Basic Help: <https://pycatia.readthedocs.io/en/latest/api/pycatia/part_interfaces/shape_factory.html>
- HybridShapeFactory extrema methods: <https://pycatia.readthedocs.io/en/latest/api/pycatia/hybrid_shape_interfaces/hybrid_shape_factory.html>
- Product `Analyze.Volume` (Product Structure, used without SPAWorkbench): <https://pycatia.readthedocs.io/en/0.3.6/api_product_structure_interfaces.html>
- Point `GetCoordinates`: <https://pycatia.readthedocs.io/en/latest/api/pycatia/hybrid_shape_interfaces/point.html>
- Products and `AddComponentsFromFiles`: <https://pycatia.readthedocs.io/en/0.8.1/api/pycatia/product_structure_interfaces/products.html>
- Viewpoint3D camera members: <https://pycatia.readthedocs.io/en/0.9.1/api/pycatia/in_interfaces/viewpoint_3d.html>
- VisPropertySet colors and opacity: <https://pycatia.readthedocs.io/en/0.6.7/api/pycatia/in_interfaces/vis_property_set.html>
- CATIA Paste Special formats and `CATPrtResultWithOutLink`: <https://www.scripting4v5.com/how-to-pastespecial-with-a-catia-macro/>
- CATIA file-selection dialog: <https://v5vb.wordpress.com/2010/02/08/file-dialogs/>

These interfaces are available in the V5 generation that includes R27. Actual execution still depends on the installed V5R27 service pack, hotfix, licenses, CATSettings, and input model health, which is why the package includes a local validation checklist.
