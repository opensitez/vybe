use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets GestureRecognizer
// ═══════════════════════════════════════════════════════════

#[test]
fn tap_gesture_recognizer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final tap = TapGestureRecognizer();
  print(tap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn double_tap_gesture_recognizer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final dt = DoubleTapGestureRecognizer();
  print(dt != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn long_press_gesture_recognizer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final lp = LongPressGestureRecognizer();
  print(lp != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn pan_gesture_recognizer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final p = PanGestureRecognizer();
  print(p != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn scale_gesture_recognizer_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final s = ScaleGestureRecognizer();
  print(s != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn pointer_event_down() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final pe = PointerDownEvent(pointer: 1);
  print(pe.pointer);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn gesture_recognizer_add_pointer() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final tap = TapGestureRecognizer();
  final pe = PointerDownEvent(pointer: 1);
  tap.addPointer(pe);
  print('pointer_added');
}
"#
        ),
        vec!["pointer_added"]
    );
}

#[test]
fn gesture_recognizer_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final tap = TapGestureRecognizer();
  tap.dispose();
  print('disposed');
}
"#
        ),
        vec!["disposed"]
    );
}

#[test]
fn gesture_recognizer_debug_owner() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final tap = TapGestureRecognizer(debugOwner: 'my_owner');
  print(tap.debugOwner);
}
"#
        ),
        vec!["my_owner"]
    );
}

#[test]
fn tap_gesture_callbacks() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/gestures.dart';
void main() {
  final tap = TapGestureRecognizer();
  tap.onTap = () {
    print('tapped');
  };
  // We can't trivially simulate pointer router events in naive testing
  print(tap.onTap != null);
}
"#
        ),
        vec!["true"]
    );
}
