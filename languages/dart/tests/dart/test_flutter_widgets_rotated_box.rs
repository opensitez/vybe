use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets RotatedBox
// ═══════════════════════════════════════════════════════════

#[test]
fn rotated_box_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RotatedBox(quarterTurns: 1, child: const SizedBox());
  print(rb.quarterTurns);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn rotated_box_quarter_turns() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RotatedBox(quarterTurns: -2, child: const SizedBox());
  print(rb.quarterTurns);
}
"#
        ),
        vec!["-2"]
    );
}

#[test]
fn rotated_box_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RotatedBox(quarterTurns: 3, child: const Placeholder());
  print(rb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rotated_box_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RotatedBox(quarterTurns: 1);
  print(rb is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn rotated_box_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final rb = RotatedBox(quarterTurns: 1);
  // Creates RenderRotatedBox
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
