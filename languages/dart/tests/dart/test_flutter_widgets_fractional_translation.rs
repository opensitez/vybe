use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets FractionalTranslation
// ═══════════════════════════════════════════════════════════

#[test]
fn fractional_translation_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ft = FractionalTranslation(
    translation: const Offset(0.5, -0.5),
    child: const SizedBox(),
  );
  print('${ft.translation.dx}:${ft.translation.dy}');
}
"#
        ),
        vec!["0.5:-0.5"]
    );
}

#[test]
fn fractional_translation_transform_hit_tests() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ft = FractionalTranslation(
    translation: const Offset(1.0, 1.0),
    transformHitTests: false,
    child: const SizedBox(),
  );
  print(ft.transformHitTests);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn fractional_translation_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ft = FractionalTranslation(
    translation: const Offset(0.0, 0.0),
    child: const Placeholder(),
  );
  print(ft.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractional_translation_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ft = FractionalTranslation(
    translation: const Offset(0.1, 0.1),
  );
  print(ft is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractional_translation_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final ft = FractionalTranslation(
    translation: const Offset(0.2, 0.2),
  );
  // Creates RenderFractionalTranslation
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
