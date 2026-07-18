use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets PhysicalModel
// ═══════════════════════════════════════════════════════════

#[test]
fn physical_model_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    child: const SizedBox(),
  );
  print(pm.color.value == 0xFF000000);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_model_shape() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    shape: BoxShape.circle,
  );
  print(pm.shape == BoxShape.circle);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_model_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    clipBehavior: Clip.antiAlias,
  );
  print(pm.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_model_elevation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    elevation: 4.0,
  );
  print(pm.elevation);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn physical_model_shadow_color() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(
    color: const Color(0xFF000000),
    shadowColor: const Color(0xFF111111),
  );
  print(pm.shadowColor.value == 0xFF111111);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_model_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final pm = PhysicalModel(color: const Color(0xFF000000));
  print(pm is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
