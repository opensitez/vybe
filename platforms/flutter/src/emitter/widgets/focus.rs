//! Focus system data types — focus nodes, the focus manager and scope nodes.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_FOCUSNODE: &[FlutterField] = &[
    FlutterField::named("debugLabel"),
    FlutterField::named_default("hasFocus", "false"),
    FlutterField::named_default("hasPrimaryFocus", "false"),
    FlutterField::named_default("canRequestFocus", "true"),
    FlutterField::named_default("skipTraversal", "false"),
    FlutterField::named_default("descendantsAreFocusable", "true"),
];

const F_FOCUSMANAGER: &[FlutterField] = &[FlutterField::named("rootScope")];
const F_FOCUSSCOPENODE: &[FlutterField] = &[FlutterField::named("debugLabel")];

pub(crate) const CLASSES: &[FlutterClass] = &[
    FlutterClass::data("FocusNode", None, F_FOCUSNODE),
    FlutterClass::data("FocusManager", None, F_FOCUSMANAGER),
    FlutterClass::data("FocusScopeNode", None, F_FOCUSSCOPENODE),
];
