//! Cupertino (iOS-style) widgets — backed by the same `vybe:gui` controls as
//! their Material counterparts.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_CUPBUTTON: &[FlutterField] = &[
    FlutterField::named("onPressed"),
    FlutterField::named("child"),
    FlutterField::named("color"),
    FlutterField::named("disabledColor"),
];

const F_CUPSLIDER: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("onChanged"),
    FlutterField::named("min"),
    FlutterField::named("max"),
    FlutterField::named("divisions"),
    FlutterField::named("activeColor"),
    FlutterField::named("thumbColor"),
];

const F_CUPSWITCH: &[FlutterField] = &[
    FlutterField::named("value"),
    FlutterField::named("onChanged"),
    FlutterField::named("activeColor"),
    FlutterField::named("trackColor"),
    FlutterField::named("thumbColor"),
];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::widget("CupertinoButton", "StatefulWidget", "Button", F_CUPBUTTON),
    FlutterClass::widget("CupertinoSlider", "StatefulWidget", "trackbar", F_CUPSLIDER),
    FlutterClass::widget("CupertinoSwitch", "StatefulWidget", "CheckBox", F_CUPSWITCH),
];
