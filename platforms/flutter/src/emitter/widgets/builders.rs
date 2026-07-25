//! Builder & inherited/implicit-animation widgets — async/state builders,
//! Hero transitions, inherited scopes, and the `Animated*` implicit family.

use crate::emitter::catalog::{FlutterClass, FlutterField, F_CHILD_ONLY};

const F_FUTUREB: &[FlutterField] = &[
    FlutterField::named("future"),
    FlutterField::named("initialData"),
    FlutterField::named("builder"),
];

const F_STREAMB: &[FlutterField] = &[
    FlutterField::named("stream"),
    FlutterField::named("builder"),
    FlutterField::named("initialData"),
];

const F_STATEFULB: &[FlutterField] = &[FlutterField::named("builder")];

const F_VLBUILDER: &[FlutterField] = &[
    FlutterField::named("valueListenable"),
    FlutterField::named("builder"),
    FlutterField::named("child"),
];

const F_ANIMBUILDER: &[FlutterField] = &[
    FlutterField::named("animation"),
    FlutterField::named("listenable"),
    FlutterField::named("child"),
    FlutterField::named("builder"),
];

const F_HERO: &[FlutterField] = &[
    FlutterField::named("tag"),
    FlutterField::named("child"),
    FlutterField::named("transitionOnUserGestures"),
    FlutterField::named("flightShuttleBuilder"),
    FlutterField::named("placeholderBuilder"),
];

const F_HEROMODE: &[FlutterField] =
    &[FlutterField::named("enabled"), FlutterField::named("child")];

const F_ICONTHEME: &[FlutterField] =
    &[FlutterField::named("data"), FlutterField::named("child")];

const F_INHNOTIFIER: &[FlutterField] =
    &[FlutterField::named("notifier"), FlutterField::named("child")];

const F_INTERACTIVE: &[FlutterField] = &[
    FlutterField::named("clipBehavior"),
    FlutterField::named("panEnabled"),
    FlutterField::named("scaleEnabled"),
    FlutterField::named("minScale"),
    FlutterField::named("maxScale"),
    FlutterField::named("constrained"),
    FlutterField::named("boundaryMargin"),
    FlutterField::named("child"),
];

const F_ANIMALIGN: &[FlutterField] = &[
    FlutterField::named("alignment"),
    FlutterField::named("duration"),
    FlutterField::named("child"),
    FlutterField::named("curve"),
];

const F_ANIMPADDING: &[FlutterField] = &[
    FlutterField::named("padding"),
    FlutterField::named("duration"),
    FlutterField::named("curve"),
    FlutterField::named("child"),
];

const F_ANIMSIZE: &[FlutterField] = &[
    FlutterField::named("duration"),
    FlutterField::named("curve"),
    FlutterField::named("child"),
];

const F_ANIMOPACITY: &[FlutterField] = &[
    FlutterField::named("opacity"),
    FlutterField::named("duration"),
    FlutterField::named("curve"),
    FlutterField::named("child"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("FutureBuilder", "StatefulWidget", "Panel", F_FUTUREB),
    FlutterClass::widget("StreamBuilder", "StatefulWidget", "Panel", F_STREAMB),
    FlutterClass::widget("StatefulBuilder", "StatefulWidget", "Panel", F_STATEFULB),
    FlutterClass::widget("ValueListenableBuilder", "StatefulWidget", "Panel", F_VLBUILDER),
    FlutterClass::widget("AnimatedBuilder", "AnimatedWidget", "Panel", F_ANIMBUILDER),
    // Scope/transition wrappers contribute no visual of their own — they carry
    // a theme, a notifier or an animation over the child, so the child realizes
    // in their place.
    FlutterClass::wrapper("Hero", "StatefulWidget", F_HERO),
    FlutterClass::wrapper("HeroMode", "StatelessWidget", F_HEROMODE),
    FlutterClass::wrapper("IconTheme", "InheritedWidget", F_ICONTHEME),
    FlutterClass::wrapper("InheritedNotifier", "InheritedWidget", F_INHNOTIFIER),
    FlutterClass::wrapper("InheritedModel", "InheritedWidget", F_CHILD_ONLY),
    FlutterClass::wrapper("AnimatedAlign", "ImplicitlyAnimatedWidget", F_ANIMALIGN),
    FlutterClass::wrapper("AnimatedPadding", "ImplicitlyAnimatedWidget", F_ANIMPADDING),
    FlutterClass::wrapper("AnimatedSize", "StatefulWidget", F_ANIMSIZE),
    FlutterClass::wrapper("AnimatedOpacity", "ImplicitlyAnimatedWidget", F_ANIMOPACITY),
    FlutterClass::widget("InteractiveViewer", "StatefulWidget", "Panel", F_INTERACTIVE),
];
