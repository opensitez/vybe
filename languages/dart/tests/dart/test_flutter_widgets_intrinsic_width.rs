use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets IntrinsicWidth
// ═══════════════════════════════════════════════════════════

#[test]
fn intrinsic_width_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iw = IntrinsicWidth();
  print(iw != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn intrinsic_width_step_width() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iw = IntrinsicWidth(stepWidth: 50.0);
  print(iw.stepWidth);
}
"#
        ),
        vec!["50.0"]
    );
}

#[test]
fn intrinsic_width_step_height() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iw = IntrinsicWidth(stepHeight: 100.0);
  print(iw.stepHeight);
}
"#
        ),
        vec!["100.0"]
    );
}

#[test]
fn intrinsic_width_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iw = IntrinsicWidth(child: const Placeholder());
  print(iw.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn intrinsic_width_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final iw = IntrinsicWidth();
  print(iw is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
