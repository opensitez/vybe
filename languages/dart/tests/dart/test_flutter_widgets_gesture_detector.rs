use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets GestureDetector
// ═══════════════════════════════════════════════════════════

#[test]
fn gesture_detector_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gd = GestureDetector(child: const SizedBox());
  print(gd is StatelessWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn gesture_detector_on_tap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gd = GestureDetector(onTap: () {}, child: const SizedBox());
  print(gd.onTap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn gesture_detector_on_double_tap() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gd = GestureDetector(onDoubleTap: () {}, child: const SizedBox());
  print(gd.onDoubleTap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn gesture_detector_on_long_press() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gd = GestureDetector(onLongPress: () {}, child: const SizedBox());
  print(gd.onLongPress != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn gesture_detector_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final gd = GestureDetector(
    behavior: HitTestBehavior.opaque,
    child: const SizedBox(),
  );
  print(gd.behavior == HitTestBehavior.opaque);
}
"#
        ),
        vec!["true"]
    );
}
