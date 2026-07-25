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

const F_VALNOTIFIER: &[FlutterField] = &[FlutterField::positional("value", 0)];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::data("AnimationController", None, F_ANIMCONTROLLER),
    FlutterClass::data("Ticker", None, F_TICKER),
    FlutterClass::data("AlwaysStoppedAnimation", None, F_ALWAYSANIM),
    FlutterClass::data("ValueNotifier", None, F_VALNOTIFIER),
];
