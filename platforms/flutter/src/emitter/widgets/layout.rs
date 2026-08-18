//! Layout & display widgets — Flex family, boxes, alignment, sizing, text.
//! Flex containers and text nodes — `Row`/`Column` are flexboxes, `Text` a
//! span.

use crate::emitter::catalog::{FlutterClass, FlutterField};

// Flex-family fields. Column/Row/Flex share everything except the main axis,
// so the shared spec is generated per direction default. These are Flutter's
// own constructor defaults, which widget code compares against
// (`Column().crossAxisAlignment == CrossAxisAlignment.center`).
macro_rules! flex_fields {
    ($direction_default:expr) => {
        &[
            FlutterField::children_list("children"),
            FlutterField::named_default("mainAxisAlignment", "MainAxisAlignment.start"),
            FlutterField::named_default("mainAxisSize", "MainAxisSize.max"),
            FlutterField::named_default("crossAxisAlignment", "CrossAxisAlignment.center"),
            FlutterField::named("textDirection"),
            FlutterField::named_default("verticalDirection", "VerticalDirection.down"),
            FlutterField::named("textBaseline"),
            FlutterField::named_default("clipBehavior", "Clip.none"),
            FlutterField::named_default("direction", $direction_default),
        ]
    };
}

// A bare `Flex` has no implied axis — it is the explicit-direction base.
const FLEX_FIELDS: &[FlutterField] = flex_fields!("Axis.horizontal");
const COLUMN_FIELDS: &[FlutterField] = flex_fields!("Axis.vertical");
const ROW_FIELDS: &[FlutterField] = flex_fields!("Axis.horizontal");

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
    // `alignment` is where the content sits across the box — CSS `text-align`,
    // the same property VCL's `Alignment` and WinForms' `TextAlign` already map
    // to. Left as its own name it wrote an `alignment=""` attribute no element
    // reads, which is why a right-aligned display sat on the left.
    FlutterField::named_role("alignment", "textalign"),
    FlutterField::named("color"),
    FlutterField::named("constraints"),
    FlutterField::named("decoration"),
    FlutterField::named("foregroundDecoration"),
    FlutterField::named("height"),
    FlutterField::named("width"),
    FlutterField::named("margin"),
    FlutterField::named("padding"),
    FlutterField::named("transform"),
    FlutterField::named_default("clipBehavior", "Clip.none"),
];

const STACK_FIELDS: &[FlutterField] = &[
    FlutterField::children_list("children"),
    FlutterField::named("alignment"),
    FlutterField::named_default("fit", "StackFit.loose"),
    FlutterField::named_default("clipBehavior", "Clip.hardEdge"),
    FlutterField::named("textDirection"),
];

const ALIGN_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_role("alignment", "textalign"),
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

// `Expanded` is a `Flexible` that forces the child to fill its share, so its
// fit is always tight; a plain `Flexible` lets the child be smaller.
const EXPANDED_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_default("flex", "1"),
    FlutterField::named_default("fit", "FlexFit.tight"),
];

const FLEXIBLE_FIELDS: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named_default("flex", "1"),
    FlutterField::named_default("fit", "FlexFit.loose"),
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
    FlutterField::named_default("fit", "BoxFit.contain"),
    FlutterField::named("alignment"),
    FlutterField::named_default("clipBehavior", "Clip.none"),
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

const F_ASPECTRATIO: &[FlutterField] = &[
    FlutterField::named("aspectRatio"),
    FlutterField::named("child"),
];

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

const F_ROTATEDBOX: &[FlutterField] = &[
    FlutterField::named("quarterTurns"),
    FlutterField::named("child"),
];

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

// ── The element each layout widget IS ───────────────────────────────────────
//
// **Flutter's layout model IS flexbox**, which is why this table is nearly
// mechanical where the WinForms one was not: `Row`/`Column` are flex containers
// in Flutter's own documentation, `Expanded` is `flex-grow`, `Padding` is
// `padding`, `SizedBox` is `width`/`height`. Naming an old control here
// threw all of that away twice over — `FlowLayoutPanel` is not an HTML tag, so
// `control_kind` matched nothing and every container rendered as a 120x20
// label.
//
// A `div` with a declared display mode is the whole implementation
// (guiplan, "the flow containers are DIVS, not custom elements"), and it means
// the ENGINE lays these out: `flex-direction`, `flex-wrap` and `flex` are
// parsed, cascaded and consumed by `vybe_widgets`, so nothing here computes a
// coordinate.
pub(crate) const CLASSES: &[FlutterClass] = &[
    // Text is phrasing content, not a box — `<span>` is what it is, and a leaf
    // that carries its own caption.
    FlutterClass::widget("Text", "StatelessWidget", "span", TEXT_FIELDS),
    FlutterClass::widget("Placeholder", "StatelessWidget", "div", &[]),
    FlutterClass::widget("Container", "StatelessWidget", "div", CONTAINER_FIELDS),
    // `Flex` is the base both directions derive from; its own default is the
    // one CSS has, `row`.
    FlutterClass::widget(
        "Flex",
        "MultiChildRenderObjectWidget",
        "div;display:flex",
        FLEX_FIELDS,
    ),
    // `flex:1` is `mainAxisSize: MainAxisSize.max`, which is a `Column`'s
    // DEFAULT — it fills its parent's main axis. Declaring it is what lets a
    // `Scaffold`'s body take the height left over by the app bar, instead of
    // shrinking to its content and leaving its `Expanded` children nothing to
    // divide.
    //
    // ⚠ It reads the PARENT's main axis, so this is the right rule only while
    // the parent is also a column — the overwhelmingly common shell. A `Column`
    // inside a `Row` will grow across instead of merely stretching down.
    // Narrowing it needs the parent's direction at the construction site.
    FlutterClass::widget(
        "Column",
        "Flex",
        "div;display:flex;flex-direction:column;flex:1",
        COLUMN_FIELDS,
    ),
    FlutterClass::widget(
        "Row",
        "Flex",
        "div;display:flex;flex-direction:row",
        ROW_FIELDS,
    ),
    // A Stack paints its children on top of one another: that is a positioned
    // container, and `Positioned` children below are `position: absolute`
    // against it. `establishes_containing_block()` already makes a `div` the
    // anchor, so the pair works with no stacking machinery of its own.
    FlutterClass::widget(
        "Stack",
        "MultiChildRenderObjectWidget",
        "div;position:relative",
        STACK_FIELDS,
    ),
    FlutterClass::widget(
        "Align",
        "SingleChildRenderObjectWidget",
        "div;display:flex",
        ALIGN_FIELDS,
    ),
    // `Center` is `Align` with both axes centred, which is exactly what these
    // two declarations say. Flutter's own definition, spelled in CSS.
    FlutterClass::widget(
        "Center",
        "Align",
        "div;display:flex;justify-content:center;align-items:center",
        CENTER_FIELDS,
    ),
    FlutterClass::widget(
        "Padding",
        "SingleChildRenderObjectWidget",
        "div",
        PADDING_FIELDS,
    ),
    FlutterClass::widget(
        "SizedBox",
        "SingleChildRenderObjectWidget",
        "div",
        SIZEDBOX_FIELDS,
    ),
    FlutterClass::widget("Icon", "StatelessWidget", "span", ICON_FIELDS),
    // `Flexible`/`Expanded` are the child's share of the parent's main axis —
    // `flex-grow`. `Expanded` is `Flexible(fit: tight)`, i.e. `flex: 1`.
    FlutterClass::widget(
        "Flexible",
        "ParentDataWidget",
        "div",
        FLEXIBLE_FIELDS,
    ),
    // **`FlexFit.tight` IS a single-cell grid.** `Expanded` forces its child
    // to fill the share it was given, on both axes — and a grid item stretches
    // to its cell on both axes by default (`align-items`/`justify-items` are
    // `stretch`). Saying it here rather than putting `height:100%` on every
    // widget that might END UP inside an `Expanded` is the difference between
    // the rule Flutter has and a guess: a `Padding` in a plain `Column` gets a
    // LOOSE constraint and must size to its child, which is why the tictactoe
    // "New Game" button came out the height of the window when `Padding` filled
    // unconditionally.
    FlutterClass::widget(
        "Expanded",
        "Flexible",
        // …with an explicit `1fr` cell on both axes. An IMPLICIT track is
        // `auto`, which sizes to its content and so grows past the share the
        // flex parent just handed out — the rows overflowed the bottom of the
        // window again. `1fr` is exactly the container, which is what a tight
        // constraint means.
        "div;flex:1;display:grid;grid-template-columns:1fr;grid-template-rows:1fr",
        EXPANDED_FIELDS,
    ),
    FlutterClass::widget(
        "Positioned",
        "ParentDataWidget",
        "div;position:absolute",
        POSITIONED_FIELDS,
    ),
    FlutterClass::widget(
        "PositionedDirectional",
        "ParentDataWidget",
        "div;position:absolute",
        F_POSDIR,
    ),
    // A Spacer takes the free space and draws nothing.
    FlutterClass::widget("Spacer", "StatelessWidget", "div;flex:1", F_SPACER),
    FlutterClass::widget("IndexedStack", "Stack", "div;position:relative", F_INDEXEDSTACK),
    // `Wrap` is the one whose name CSS shares outright.
    FlutterClass::widget(
        "Wrap",
        "MultiChildRenderObjectWidget",
        "div;display:flex;flex-wrap:wrap",
        F_WRAP,
    ),
    // Sizing/transform wrappers: the effect (fit, fraction, rotation, matrix)
    // has no backing-control command, so they realize their child directly
    // rather than wrapping it in an inert Panel.
    FlutterClass::wrapper("FittedBox", "SingleChildRenderObjectWidget", F_FITTEDBOX),
    FlutterClass::wrapper(
        "FractionalTranslation",
        "SingleChildRenderObjectWidget",
        F_FRACTRANS,
    ),
    FlutterClass::wrapper(
        "FractionallySizedBox",
        "SingleChildRenderObjectWidget",
        F_FRACSIZED,
    ),
    FlutterClass::wrapper(
        "IntrinsicHeight",
        "SingleChildRenderObjectWidget",
        F_CHILD_ONLY,
    ),
    FlutterClass::wrapper(
        "IntrinsicWidth",
        "SingleChildRenderObjectWidget",
        F_INTRINSICW,
    ),
    FlutterClass::wrapper(
        "AspectRatio",
        "SingleChildRenderObjectWidget",
        F_ASPECTRATIO,
    ),
    FlutterClass::wrapper("Baseline", "SingleChildRenderObjectWidget", F_BASELINE),
    FlutterClass::wrapper("Transform", "SingleChildRenderObjectWidget", F_TRANSFORM),
    FlutterClass::wrapper("RotatedBox", "SingleChildRenderObjectWidget", F_ROTATEDBOX),
    FlutterClass::wrapper("DefaultTextStyle", "StatelessWidget", F_DEFTEXTSTYLE),
    FlutterClass::widget(
        "RichText",
        "MultiChildRenderObjectWidget",
        "span",
        F_RICHTEXT,
    ),
];
