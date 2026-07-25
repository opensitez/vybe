//! Layout & display widgets — Flex family, boxes, alignment, sizing, text.
//! Backed by `vybe:gui` flow-layout panels and labels.

use crate::emitter::catalog::{FlutterClass, FlutterField};

// Flex-family fields (Column/Row/Flex share these). `children` defaults to an
// empty list; enum-valued fields default to null until the enum surface lands.
const FLEX_FIELDS: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("mainAxisAlignment"),
    FlutterField::named("mainAxisSize"),
    FlutterField::named("crossAxisAlignment"),
    FlutterField::named("textDirection"),
    FlutterField::named("verticalDirection"),
    FlutterField::named("textBaseline"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("direction"),
];

const TEXT_FIELDS: &[FlutterField] = &[
    FlutterField::positional("data", 0),
    FlutterField::named("style"),
    FlutterField::named("textAlign"),
    FlutterField::named("textDirection"),
    FlutterField::named_default("softWrap", "true"),
    FlutterField::named("overflow"),
    FlutterField::named("maxLines"),
    FlutterField::named("textSpan"),
];

const CONTAINER_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("alignment"),
    FlutterField::named("color"),
    FlutterField::named("constraints"),
    FlutterField::named("decoration"),
    FlutterField::named("foregroundDecoration"),
    FlutterField::named("height"),
    FlutterField::named("width"),
    FlutterField::named("margin"),
    FlutterField::named("padding"),
    FlutterField::named("transform"),
    FlutterField::named("clipBehavior"),
];

const STACK_FIELDS: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("alignment"),
    FlutterField::named("fit"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("textDirection"),
];

const ALIGN_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("alignment"),
    FlutterField::named("heightFactor"),
    FlutterField::named("widthFactor"),
];

const CENTER_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("heightFactor"),
    FlutterField::named("widthFactor"),
];

const PADDING_FIELDS: &[FlutterField] =
    &[FlutterField::named("padding"), FlutterField::named("child")];

const SIZEDBOX_FIELDS: &[FlutterField] = &[
    FlutterField::named("width"),
    FlutterField::named("height"),
    FlutterField::named("child"),
];

const ICON_FIELDS: &[FlutterField] = &[
    FlutterField::positional("icon", 0),
    FlutterField::named("color"),
    FlutterField::named("size"),
    FlutterField::named("semanticLabel"),
    FlutterField::named("textDirection"),
];

const EXPANDED_FIELDS: &[FlutterField] =
    &[FlutterField::named("child"), FlutterField::named_default("flex", "1")];

const FLEXIBLE_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_default("flex", "1"),
    FlutterField::named("fit"),
];

const POSITIONED_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
    FlutterField::named("width"),
    FlutterField::named("height"),
];

const F_POSDIR: &[FlutterField] = &[
    FlutterField::named("start"),
    FlutterField::named("end"),
    FlutterField::named("top"),
    FlutterField::named("bottom"),
    FlutterField::named("child"),
];

const F_SPACER: &[FlutterField] = &[FlutterField::named_default("flex", "1")];

const F_INDEXEDSTACK: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("index"),
    FlutterField::named("alignment"),
    FlutterField::named("sizing"),
    FlutterField::named("textDirection"),
];

const F_WRAP: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("direction"),
    FlutterField::named("alignment"),
    FlutterField::named("spacing"),
    FlutterField::named("runAlignment"),
    FlutterField::named("runSpacing"),
    FlutterField::named("crossAxisAlignment"),
    FlutterField::named("textDirection"),
    FlutterField::named("verticalDirection"),
    FlutterField::named("clipBehavior"),
];

const F_FITTEDBOX: &[FlutterField] = &[
    FlutterField::named("fit"),
    FlutterField::named("alignment"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("child"),
];

const F_FRACTRANS: &[FlutterField] = &[
    FlutterField::named("translation"),
    FlutterField::named("transformHitTests"),
    FlutterField::named("child"),
];

const F_FRACSIZED: &[FlutterField] = &[
    FlutterField::named("widthFactor"),
    FlutterField::named("heightFactor"),
    FlutterField::named("alignment"),
    FlutterField::named("child"),
];

const F_INTRINSICW: &[FlutterField] = &[
    FlutterField::named("stepWidth"),
    FlutterField::named("stepHeight"),
    FlutterField::named("child"),
];

const F_ASPECTRATIO: &[FlutterField] =
    &[FlutterField::named("aspectRatio"), FlutterField::named("child")];

const F_BASELINE: &[FlutterField] = &[
    FlutterField::named("baseline"),
    FlutterField::named("baselineType"),
    FlutterField::named("child"),
];

const F_TRANSFORM: &[FlutterField] = &[
    FlutterField::named("transform"),
    FlutterField::named("origin"),
    FlutterField::named("alignment"),
    FlutterField::named("transformHitTests"),
    FlutterField::named("angle"),
    FlutterField::named("offset"),
    FlutterField::named("scale"),
    FlutterField::named("scaleX"),
    FlutterField::named("scaleY"),
    FlutterField::named("child"),
];

const F_ROTATEDBOX: &[FlutterField] =
    &[FlutterField::named("quarterTurns"), FlutterField::named("child")];

const F_RICHTEXT: &[FlutterField] = &[
    FlutterField::named("text"),
    FlutterField::named("textAlign"),
    FlutterField::named("textDirection"),
    FlutterField::named("softWrap"),
    FlutterField::named("overflow"),
    FlutterField::named("maxLines"),
];

const F_DEFTEXTSTYLE: &[FlutterField] =
    &[FlutterField::named("style"), FlutterField::named("child")];

use crate::emitter::catalog::F_CHILD_ONLY;

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("Text", "StatelessWidget", "Label", TEXT_FIELDS),
    FlutterClass::widget("Placeholder", "StatelessWidget", "Panel", &[]),
    FlutterClass::widget("Container", "StatelessWidget", "FlowLayoutPanel", CONTAINER_FIELDS),
    FlutterClass::widget("Flex", "MultiChildRenderObjectWidget", "FlowLayoutPanel", FLEX_FIELDS),
    FlutterClass::widget("Column", "Flex", "FlowLayoutPanel", FLEX_FIELDS),
    FlutterClass::widget("Row", "Flex", "HFlowLayoutPanel", FLEX_FIELDS),
    FlutterClass::widget("Stack", "MultiChildRenderObjectWidget", "FlowLayoutPanel", STACK_FIELDS),
    FlutterClass::widget("Align", "SingleChildRenderObjectWidget", "FlowLayoutPanel", ALIGN_FIELDS),
    FlutterClass::widget("Center", "Align", "FlowLayoutPanel", CENTER_FIELDS),
    FlutterClass::widget("Padding", "SingleChildRenderObjectWidget", "FlowLayoutPanel", PADDING_FIELDS),
    FlutterClass::widget("SizedBox", "SingleChildRenderObjectWidget", "FlowLayoutPanel", SIZEDBOX_FIELDS),
    FlutterClass::widget("Icon", "StatelessWidget", "Label", ICON_FIELDS),
    FlutterClass::widget("Flexible", "ParentDataWidget", "FlowLayoutPanel", FLEXIBLE_FIELDS),
    FlutterClass::widget("Expanded", "Flexible", "FlowLayoutPanel", EXPANDED_FIELDS),
    FlutterClass::widget("Positioned", "ParentDataWidget", "Panel", POSITIONED_FIELDS),
    FlutterClass::widget("PositionedDirectional", "ParentDataWidget", "Panel", F_POSDIR),
    FlutterClass::widget("Spacer", "StatelessWidget", "Panel", F_SPACER),
    FlutterClass::widget("IndexedStack", "Stack", "FlowLayoutPanel", F_INDEXEDSTACK),
    FlutterClass::widget("Wrap", "MultiChildRenderObjectWidget", "FlowLayoutPanel", F_WRAP),
    // Sizing/transform wrappers: the effect (fit, fraction, rotation, matrix)
    // has no backing-control command, so they realize their child directly
    // rather than wrapping it in an inert Panel.
    FlutterClass::wrapper("FittedBox", "SingleChildRenderObjectWidget", F_FITTEDBOX),
    FlutterClass::wrapper("FractionalTranslation", "SingleChildRenderObjectWidget", F_FRACTRANS),
    FlutterClass::wrapper("FractionallySizedBox", "SingleChildRenderObjectWidget", F_FRACSIZED),
    FlutterClass::wrapper("IntrinsicHeight", "SingleChildRenderObjectWidget", F_CHILD_ONLY),
    FlutterClass::wrapper("IntrinsicWidth", "SingleChildRenderObjectWidget", F_INTRINSICW),
    FlutterClass::wrapper("AspectRatio", "SingleChildRenderObjectWidget", F_ASPECTRATIO),
    FlutterClass::wrapper("Baseline", "SingleChildRenderObjectWidget", F_BASELINE),
    FlutterClass::wrapper("Transform", "SingleChildRenderObjectWidget", F_TRANSFORM),
    FlutterClass::wrapper("RotatedBox", "SingleChildRenderObjectWidget", F_ROTATEDBOX),
    FlutterClass::wrapper("DefaultTextStyle", "StatelessWidget", F_DEFTEXTSTYLE),
    FlutterClass::widget("RichText", "MultiChildRenderObjectWidget", "Label", F_RICHTEXT),
];
