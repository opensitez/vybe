use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets RepaintBoundary
// ═══════════════════════════════════════════════════════════

#[test]
fn repaint_boundary_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RepaintBoundary(child: const SizedBox());
  print(rb != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn repaint_boundary_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RepaintBoundary(child: const Placeholder());
  print(rb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn repaint_boundary_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final rb = RepaintBoundary();
  print(rb is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn repaint_boundary_wrap_static() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = RepaintBoundary.wrap(const SizedBox(), 1);
  print(w is RepaintBoundary);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn repaint_boundary_wrap_all_static() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final list = RepaintBoundary.wrapAll([const SizedBox(), const SizedBox()]);
  print(list.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn repaint_boundary_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final rb = RepaintBoundary();
  // Creates RenderRepaintBoundary
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
