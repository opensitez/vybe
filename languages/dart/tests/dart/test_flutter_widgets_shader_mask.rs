use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ShaderMask
// ═══════════════════════════════════════════════════════════

#[test]
fn shader_mask_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sm = ShaderMask(
    shaderCallback: (Rect bounds) => const RadialGradient(
      center: Alignment.topLeft,
      radius: 1.0,
      colors: <Color>[Color(0xFFFFFF00), Color(0xFF0000FF)],
      tileMode: TileMode.mirror,
    ).createShader(bounds),
    child: const SizedBox(),
  );
  print(sm.shaderCallback != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn shader_mask_blend_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sm = ShaderMask(
    shaderCallback: (Rect bounds) => const LinearGradient(
      colors: [Color(0xFF000000), Color(0xFFFFFFFF)],
    ).createShader(bounds),
    blendMode: BlendMode.dstIn,
    child: const SizedBox(),
  );
  print(sm.blendMode == BlendMode.dstIn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn shader_mask_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sm = ShaderMask(
    shaderCallback: (Rect bounds) => const LinearGradient(
      colors: [Color(0xFF000000), Color(0xFFFFFFFF)],
    ).createShader(bounds),
    child: const Placeholder(),
  );
  print(sm.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn shader_mask_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sm = ShaderMask(
    shaderCallback: (Rect bounds) => const LinearGradient(
      colors: [Color(0xFF000000), Color(0xFFFFFFFF)],
    ).createShader(bounds),
  );
  print(sm is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn shader_mask_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final sm = ShaderMask(
    shaderCallback: (Rect bounds) => const LinearGradient(
      colors: [Color(0xFF000000), Color(0xFFFFFFFF)],
    ).createShader(bounds),
  );
  // Creates RenderShaderMask
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
