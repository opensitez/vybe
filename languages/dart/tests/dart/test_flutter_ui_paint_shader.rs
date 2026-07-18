use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: dart:ui Paint & Shader
// ═══════════════════════════════════════════════════════════

#[test]
fn paint_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  print(paint != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_color() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.color = Color(0xFF123456);
  print(paint.color.value);
}
"#
        ),
        vec!["4279383126"] // 0xFF123456
    );
}

#[test]
fn paint_is_anti_alias() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  // Default is true in Flutter
  print(paint.isAntiAlias);
  paint.isAntiAlias = false;
  print(paint.isAntiAlias);
}
"#
        ),
        vec!["true\nfalse"]
    );
}

#[test]
fn paint_blend_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.blendMode = BlendMode.multiply;
  print(paint.blendMode == BlendMode.multiply);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_style() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.style = PaintingStyle.stroke;
  print(paint.style == PaintingStyle.stroke);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_stroke_width() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.strokeWidth = 5.5;
  print(paint.strokeWidth);
}
"#
        ),
        vec!["5.5"]
    );
}

#[test]
fn paint_stroke_cap() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.strokeCap = StrokeCap.round;
  print(paint.strokeCap == StrokeCap.round);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_stroke_join() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.strokeJoin = StrokeJoin.bevel;
  print(paint.strokeJoin == StrokeJoin.bevel);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_stroke_miter_limit() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.strokeMiterLimit = 2.0;
  print(paint.strokeMiterLimit);
}
"#
        ),
        vec!["2.0"]
    );
}

#[test]
fn paint_mask_filter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.maskFilter = MaskFilter.blur(BlurStyle.normal, 3.0);
  print(paint.maskFilter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_color_filter() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.colorFilter = ColorFilter.mode(Color(0xFF000000), BlendMode.srcOver);
  print(paint.colorFilter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn color_filter_matrix() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final filter = ColorFilter.matrix([
    1, 0, 0, 0, 0,
    0, 1, 0, 0, 0,
    0, 0, 1, 0, 0,
    0, 0, 0, 1, 0,
  ]);
  print(filter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_shader_linear_gradient() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.shader = Gradient.linear(
    Offset(0, 0),
    Offset(10, 10),
    [Color(0xFF000000), Color(0xFFFFFFFF)],
  );
  print(paint.shader != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_shader_radial_gradient() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.shader = Gradient.radial(
    Offset(5, 5),
    10.0,
    [Color(0xFF000000), Color(0xFFFFFFFF)],
  );
  print(paint.shader != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_shader_sweep_gradient() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.shader = Gradient.sweep(
    Offset(5, 5),
    [Color(0xFF000000), Color(0xFFFFFFFF)],
  );
  print(paint.shader != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_image_filter_blur() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.imageFilter = ImageFilter.blur(sigmaX: 5.0, sigmaY: 5.0);
  print(paint.imageFilter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_image_filter_matrix() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
import 'dart:typed_data';
void main() {
  final paint = Paint();
  final matrix = Float64List(16);
  matrix[0] = 1.0; matrix[5] = 1.0; matrix[10] = 1.0; matrix[15] = 1.0;
  paint.imageFilter = ImageFilter.matrix(matrix);
  print(paint.imageFilter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_invert_colors() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint();
  paint.invertColors = true;
  print(paint.invertColors);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fragment_shader_compiles() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  // We can't actually compile a shader from asset without mock host logic,
  // but we can check FragmentProgram type exists.
  print(FragmentProgram != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_chaining_cascades() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint()
    ..color = Color(0xFFFFFFFF)
    ..style = PaintingStyle.stroke
    ..strokeWidth = 2.0;
  print('${paint.strokeWidth}:${paint.color.value}');
}
"#
        ),
        vec!["2.0:4294967295"]
    );
}
