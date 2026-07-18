use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets IntrinsicHeight
// ═══════════════════════════════════════════════════════════

#[test]
fn intrinsic_height_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ih = IntrinsicHeight();
  print(ih != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn intrinsic_height_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ih = IntrinsicHeight(child: const Placeholder());
  print(ih.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn intrinsic_height_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ih = IntrinsicHeight();
  print(ih is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn intrinsic_height_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final ih = IntrinsicHeight();
  // Creates RenderIntrinsicHeight
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
