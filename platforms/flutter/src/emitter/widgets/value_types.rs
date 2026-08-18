//! Value types — colors, geometry (Offset/Size), text styling, decorations,
//! gradients, semantics and other pure-data types with no backing control.
//! Constructed with named/positional args and read back by field.

use crate::emitter::catalog::{FlutterClass, FlutterField};

// `Color(0xFF00FF00)` — the packed ARGB int is the positional field; the four
// channels are derived from it. Flutter exposes them as getters, but a Color is
// immutable, so the walker's construction desugar computes them once (the same
// treatment `Rect` gets for `width`/`height`).
const COLOR_FIELDS: &[FlutterField] = &[
    FlutterField::positional("value", 0),
    FlutterField::named("alpha"),
    FlutterField::named("red"),
    FlutterField::named("green"),
    FlutterField::named("blue"),
];

// EdgeInsets carries four resolved edge insets; its named constructors
// (`all`/`symmetric`/`only`/`fromLTRB`) are desugared in the Dart walker to
// this four-field construction.
const F_EDGEINSETS: &[FlutterField] = &[
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
];

// `Alignment(x, y)` on Flutter's own -1..1 axes.
const F_ALIGNMENT: &[FlutterField] = &[
    FlutterField::positional("x", 0),
    FlutterField::positional("y", 1),
];

const F_OFFSET: &[FlutterField] = &[
    FlutterField::positional("dx", 0),
    FlutterField::positional("dy", 1),
];

// A `Rect` stores its four edges; `width`/`height`/`center` are derived in
// Flutter, so the walker's named-constructor desugar computes them once at
// construction (they can never drift — a Rect is immutable).
const F_RECT: &[FlutterField] = &[
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
    FlutterField::named("width"),
    FlutterField::named("height"),
];

// `Radius.circular(r)` is an elliptical radius with equal axes.
const F_RADIUS: &[FlutterField] = &[FlutterField::named("x"), FlutterField::named("y")];

const F_RRECT: &[FlutterField] = &[
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
    FlutterField::named("width"),
    FlutterField::named("height"),
    FlutterField::named("tlRadius"),
    FlutterField::named("trRadius"),
    FlutterField::named("blRadius"),
    FlutterField::named("brRadius"),
    // Each corner's radius is also readable as its two scalar axes.
    FlutterField::named("tlRadiusX"),
    FlutterField::named("tlRadiusY"),
    FlutterField::named("trRadiusX"),
    FlutterField::named("trRadiusY"),
    FlutterField::named("blRadiusX"),
    FlutterField::named("blRadiusY"),
    FlutterField::named("brRadiusX"),
    FlutterField::named("brRadiusY"),
];

const F_BOXCONSTRAINTS: &[FlutterField] = &[
    FlutterField::named_default("minWidth", "0"),
    FlutterField::named_default("maxWidth", "double.infinity"),
    FlutterField::named_default("minHeight", "0"),
    FlutterField::named_default("maxHeight", "double.infinity"),
];

const F_RELATIVERECT: &[FlutterField] = &[
    FlutterField::named("left"),
    FlutterField::named("top"),
    FlutterField::named("right"),
    FlutterField::named("bottom"),
];
const F_SIZE: &[FlutterField] = &[
    FlutterField::positional("width", 0),
    FlutterField::positional("height", 1),
];
const F_ICONDATA: &[FlutterField] = &[
    FlutterField::positional("codePoint", 0),
    FlutterField::named("fontFamily"),
];
const F_ICONTHEMEDATA: &[FlutterField] =
    &[FlutterField::named("color"), FlutterField::named("size")];
const F_TEXTSTYLE: &[FlutterField] = &[
    FlutterField::named("fontSize"),
    FlutterField::named("color"),
    FlutterField::named("fontWeight"),
    FlutterField::named("fontFamily"),
];
const F_TEXTSPAN2: &[FlutterField] = &[
    FlutterField::named("text"),
    FlutterField::named("children"),
    FlutterField::named("style"),
];
const F_BORDERSIDE: &[FlutterField] = &[FlutterField::named("color"), FlutterField::named("width")];
const F_BOXDECORATION: &[FlutterField] = &[
    FlutterField::named("color"),
    FlutterField::named("shape"),
    FlutterField::named("borderRadius"),
    FlutterField::named("border"),
    FlutterField::named("gradient"),
];
const F_LINEARGRAD: &[FlutterField] = &[
    FlutterField::named("colors"),
    FlutterField::named("begin"),
    FlutterField::named("end"),
    FlutterField::named("stops"),
    FlutterField::named("tileMode"),
];
const F_RADIALGRAD: &[FlutterField] = &[
    FlutterField::named("center"),
    FlutterField::named("radius"),
    FlutterField::named("colors"),
    FlutterField::named("tileMode"),
];
const F_TIMEOFDAY: &[FlutterField] = &[FlutterField::named("hour"), FlutterField::named("minute")];
const F_ROUTESETTINGS: &[FlutterField] = &[
    FlutterField::named("name"),
    FlutterField::named("arguments"),
];
const F_SEMNODE: &[FlutterField] = &[
    FlutterField::named("label"),
    FlutterField::named("value"),
    FlutterField::named("hint"),
    FlutterField::named("increasedValue"),
    FlutterField::named("decreasedValue"),
    FlutterField::named("hasCheckedState"),
    FlutterField::named("isChecked"),
    FlutterField::named("rect"),
    FlutterField::named("tags"),
];
const F_LABEL_ONLY: &[FlutterField] = &[FlutterField::named("label")];
const F_SEMACTION: &[FlutterField] = &[FlutterField::named("hint"), FlutterField::named("action")];
const F_SEMTAG: &[FlutterField] = &[FlutterField::positional("name", 0)];
const F_FITTEDSIZES: &[FlutterField] = &[
    FlutterField::positional("source", 0),
    FlutterField::positional("destination", 1),
];
const F_DIAGBUILDER: &[FlutterField] = &[FlutterField::named("properties")];
const F_ASYNCSNAPSHOT: &[FlutterField] = &[
    FlutterField::named("connectionState"),
    FlutterField::named("hasData"),
    FlutterField::named("data"),
    FlutterField::named("hasError"),
];
const F_IMGCONFIG: &[FlutterField] = &[
    FlutterField::named("size"),
    FlutterField::named("devicePixelRatio"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::data("Color", None, COLOR_FIELDS),
    FlutterClass::data("EdgeInsets", None, F_EDGEINSETS),
    FlutterClass::data("EdgeInsetsDirectional", None, F_EDGEINSETS),
    // A value type like the rest — and it has to be one HERE, not only a class
    // in `runtime.dart`, or its `toString` is never dispatched and the
    // alignment reaches CSS as the literal `[object]`.
    FlutterClass::data("Alignment", None, F_ALIGNMENT),
    FlutterClass::data("Offset", None, F_OFFSET),
    FlutterClass::data("Rect", None, F_RECT),
    FlutterClass::data("Radius", None, F_RADIUS),
    FlutterClass::data("RRect", None, F_RRECT),
    FlutterClass::data("RelativeRect", None, F_RELATIVERECT),
    FlutterClass::data("BoxConstraints", None, F_BOXCONSTRAINTS),
    FlutterClass::data("Size", None, F_SIZE),
    FlutterClass::data("FractionalOffset", None, F_OFFSET),
    FlutterClass::data("IconData", None, F_ICONDATA),
    FlutterClass::data("IconThemeData", None, F_ICONTHEMEDATA),
    FlutterClass::data("TextStyle", None, F_TEXTSTYLE),
    FlutterClass::data("TextSpan", None, F_TEXTSPAN2),
    FlutterClass::data("BorderSide", None, F_BORDERSIDE),
    FlutterClass::data("BoxDecoration", None, F_BOXDECORATION),
    FlutterClass::data("LinearGradient", None, F_LINEARGRAD),
    FlutterClass::data("RadialGradient", None, F_RADIALGRAD),
    FlutterClass::data("TimeOfDay", None, F_TIMEOFDAY),
    FlutterClass::data("RouteSettings", None, F_ROUTESETTINGS),
    FlutterClass::data("SemanticsNode", None, F_SEMNODE),
    FlutterClass::data("SemanticsConfiguration", None, F_LABEL_ONLY),
    FlutterClass::data("SemanticsProperties", None, F_LABEL_ONLY),
    FlutterClass::data("CustomSemanticsAction", None, F_SEMACTION),
    FlutterClass::data("SemanticsTag", None, F_SEMTAG),
    FlutterClass::data("FittedSizes", None, F_FITTEDSIZES),
    FlutterClass::data("DiagnosticPropertiesBuilder", None, F_DIAGBUILDER),
    FlutterClass::data("AsyncSnapshot", None, F_ASYNCSNAPSHOT),
    FlutterClass::data("ImageConfiguration", None, F_IMGCONFIG),
];
