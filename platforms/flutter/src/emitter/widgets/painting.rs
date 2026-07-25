//! Painting / effect widgets — opacity, clipping, decoration, custom paint,
//! shaders and physical-model shadows. All lower onto a `vybe:gui` Panel.

use crate::emitter::catalog::{FlutterClass, FlutterField, F_CHILD_ONLY};

const F_OPACITY: &[FlutterField] = &[
    FlutterField::named("opacity"),
    FlutterField::named("alwaysIncludeSemantics"),
    FlutterField::named("child"),
];

const F_SLIVEROPACITY: &[FlutterField] = &[
    FlutterField::named("opacity"),
    FlutterField::named("alwaysIncludeSemantics"),
    FlutterField::named("sliver"),
];

const F_DECORATEDBOX: &[FlutterField] = &[
    FlutterField::named("decoration"),
    FlutterField::named("position"),
    FlutterField::named("child"),
];

const F_CLIPPER: &[FlutterField] = &[
    FlutterField::named("clipBehavior"),
    FlutterField::named("child"),
    FlutterField::named("clipper"),
];

const F_CLIPRRECT: &[FlutterField] = &[
    FlutterField::named("borderRadius"),
    FlutterField::named("child"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("clipper"),
];

const F_COLORFILTERED: &[FlutterField] =
    &[FlutterField::named("colorFilter"), FlutterField::named("child")];

const F_CUSTOMPAINT: &[FlutterField] = &[
    FlutterField::named("painter"),
    FlutterField::named("foregroundPainter"),
    FlutterField::named("size"),
    FlutterField::named("isComplex"),
    FlutterField::named("willChange"),
    FlutterField::named("child"),
];

const F_BACKDROP: &[FlutterField] = &[
    FlutterField::named("filter"),
    FlutterField::named("blendMode"),
    FlutterField::named("child"),
];

const F_SHADERMASK: &[FlutterField] = &[
    FlutterField::named("shaderCallback"),
    FlutterField::named("blendMode"),
    FlutterField::named("child"),
];

const F_PHYSMODEL: &[FlutterField] = &[
    FlutterField::named("color"),
    FlutterField::named("child"),
    FlutterField::named("shape"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("elevation"),
    FlutterField::named("shadowColor"),
];

const F_PHYSSHAPE: &[FlutterField] = &[
    FlutterField::named("clipper"),
    FlutterField::named("color"),
    FlutterField::named("child"),
    FlutterField::named("clipBehavior"),
    FlutterField::named("elevation"),
    FlutterField::named("shadowColor"),
];

// Every widget here is a pure paint EFFECT over its child (opacity, clip,
// colour filter, shadow). None of those effects is expressible on the backing
// `vybe:gui` controls, so each realizes transparently — the child renders in
// the wrapper's place rather than inside an inert Panel.
pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::wrapper("Opacity", "SingleChildRenderObjectWidget", F_OPACITY),
    FlutterClass::wrapper("SliverOpacity", "SingleChildRenderObjectWidget", F_SLIVEROPACITY),
    FlutterClass::wrapper("DecoratedBox", "SingleChildRenderObjectWidget", F_DECORATEDBOX),
    FlutterClass::wrapper("ClipOval", "SingleChildRenderObjectWidget", F_CLIPPER),
    FlutterClass::wrapper("ClipPath", "SingleChildRenderObjectWidget", F_CLIPPER),
    FlutterClass::wrapper("ClipRect", "SingleChildRenderObjectWidget", F_CLIPPER),
    FlutterClass::wrapper("ClipRRect", "SingleChildRenderObjectWidget", F_CLIPRRECT),
    FlutterClass::wrapper("ColorFiltered", "SingleChildRenderObjectWidget", F_COLORFILTERED),
    FlutterClass::wrapper("BackdropFilter", "SingleChildRenderObjectWidget", F_BACKDROP),
    FlutterClass::wrapper("ShaderMask", "SingleChildRenderObjectWidget", F_SHADERMASK),
    FlutterClass::wrapper("PhysicalModel", "SingleChildRenderObjectWidget", F_PHYSMODEL),
    FlutterClass::wrapper("PhysicalShape", "SingleChildRenderObjectWidget", F_PHYSSHAPE),
    FlutterClass::wrapper("RepaintBoundary", "SingleChildRenderObjectWidget", F_CHILD_ONLY),
    // CustomPaint owns a painter callback rather than a child effect — it keeps
    // a real Panel so the painted surface has somewhere to live.
    FlutterClass::widget("CustomPaint", "SingleChildRenderObjectWidget", "Panel", F_CUSTOMPAINT),
];
