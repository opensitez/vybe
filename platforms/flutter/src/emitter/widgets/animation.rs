//! Animation data types — controllers, tickers, notifiers and constant
//! animations. Pure data (no backing control); tests construct and read them.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_ANIMCONTROLLER: &[FlutterField] = &[
    FlutterField::named("vsync"),
    FlutterField::named("duration"),
    FlutterField::named("lowerBound"),
    FlutterField::named("upperBound"),
];

const F_TICKER: &[FlutterField] = &[FlutterField::positional("onTick", 0)];

const F_ALWAYSANIM: &[FlutterField] = &[FlutterField::positional("value", 0)];


pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::data("AnimationController", None, F_ANIMCONTROLLER),
    FlutterClass::data("Ticker", None, F_TICKER),
    FlutterClass::data("AlwaysStoppedAnimation", None, F_ALWAYSANIM),
    // `ValueNotifier` is an ADAPTER CLASS now
    // (`core_classes/value_notifier.rs`) — it needs real behaviour
    // (`addListener`/`notifyListeners`, a setter that compares with `==`), and
    // a catalog row can only describe fields. Leaving the row here as well
    // meant construction ran the CATALOG ctor, which filled a `value` field the
    // class's own property never reads, so every notifier was born null.
];
