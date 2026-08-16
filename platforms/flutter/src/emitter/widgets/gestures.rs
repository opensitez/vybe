//! Gesture & pointer widgets — detectors, ink responses, drag/drop, dismiss,
//! and the gesture-recognizer / pointer-event data types.

use crate::emitter::catalog::{FlutterClass, FlutterField};

const F_GESTURE: &[FlutterField] = &[
    FlutterField::named("onTap"),
    FlutterField::named("onDoubleTap"),
    FlutterField::named("onLongPress"),
    FlutterField::named("behavior"),
    FlutterField::named("child"),
];

const F_INKWELL: &[FlutterField] = &[
    FlutterField::named("child"),
    FlutterField::named("onTap"),
    FlutterField::named("onDoubleTap"),
    FlutterField::named("onLongPress"),
    FlutterField::named("splashColor"),
    FlutterField::named("highlightColor"),
    FlutterField::named("borderRadius"),
];

const F_DRAGGABLE: &[FlutterField] = &[
    FlutterField::named("data"),
    FlutterField::named("feedback"),
    FlutterField::named("child"),
    FlutterField::named("childWhenDragging"),
];

const F_DRAGTARGET: &[FlutterField] = &[
    FlutterField::named("builder"),
    FlutterField::named("onWillAcceptWithDetails"),
    FlutterField::named("onAcceptWithDetails"),
];

const F_DISMISSIBLE: &[FlutterField] = &[
    FlutterField::named("key"),
    FlutterField::named("child"),
    FlutterField::named("background"),
    FlutterField::named("secondaryBackground"),
    FlutterField::named("direction"),
    FlutterField::named("onDismissed"),
];

const F_RECOGNIZER: &[FlutterField] = &[
    FlutterField::named("debugOwner"),
    FlutterField::named("onTap"),
];

const F_POINTER: &[FlutterField] = &[FlutterField::named("pointer")];

pub(crate) const CLASSES: &[FlutterClass] = &[
    // A gesture wrapper is a plain box that listens. The listening is the
    // `on*` fields; the box is a `div`.
    FlutterClass::widget("GestureDetector", "StatelessWidget", "div", F_GESTURE),
    FlutterClass::widget("InkWell", "StatelessWidget", "div", F_INKWELL),
    FlutterClass::widget("Draggable", "StatefulWidget", "div", F_DRAGGABLE),
    FlutterClass::widget("DragTarget", "StatefulWidget", "div", F_DRAGTARGET),
    FlutterClass::widget("Dismissible", "StatefulWidget", "div", F_DISMISSIBLE),
    FlutterClass::data("TapGestureRecognizer", None, F_RECOGNIZER),
    FlutterClass::data("DoubleTapGestureRecognizer", None, F_RECOGNIZER),
    FlutterClass::data("LongPressGestureRecognizer", None, F_RECOGNIZER),
    FlutterClass::data("PanGestureRecognizer", None, F_RECOGNIZER),
    FlutterClass::data("ScaleGestureRecognizer", None, F_RECOGNIZER),
    FlutterClass::data("PointerDownEvent", None, F_POINTER),
];
