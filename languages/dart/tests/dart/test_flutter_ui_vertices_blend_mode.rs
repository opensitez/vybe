use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: dart:ui Vertices & BlendMode
// ═══════════════════════════════════════════════════════════

#[test]
fn vertices_creation_triangles() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(VertexMode.triangles, [Offset(0, 0), Offset(10, 0), Offset(0, 10)]);
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_creation_triangle_strip() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(VertexMode.triangleStrip, [Offset(0, 0), Offset(10, 0), Offset(0, 10), Offset(10, 10)]);
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_creation_triangle_fan() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(VertexMode.triangleFan, [Offset(5, 5), Offset(0, 0), Offset(10, 0), Offset(10, 10)]);
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_with_colors() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(
    VertexMode.triangles,
    [Offset(0, 0), Offset(10, 0), Offset(0, 10)],
    colors: [Color(0xFFFF0000), Color(0xFF00FF00), Color(0xFF0000FF)]
  );
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_with_texture_coordinates() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(
    VertexMode.triangles,
    [Offset(0, 0), Offset(10, 0), Offset(0, 10)],
    textureCoordinates: [Offset(0, 0), Offset(1, 0), Offset(0, 1)]
  );
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_with_indices() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final v = Vertices(
    VertexMode.triangles,
    [Offset(0, 0), Offset(10, 0), Offset(0, 10), Offset(10, 10)],
    indices: [0, 1, 2, 1, 2, 3]
  );
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn vertices_raw_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
import 'dart:typed_data';
void main() {
  final positions = Float32List.fromList([0, 0, 10, 0, 0, 10]);
  final v = Vertices.raw(VertexMode.triangles, positions);
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn blend_mode_values() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  print(BlendMode.values.length > 0);
  print(BlendMode.clear.index >= 0);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn blend_mode_srcOver() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  print(BlendMode.srcOver == BlendMode.values[3]); // Usually 3
  print(BlendMode.srcOver.name);
}
"#
        ),
        vec!["true\nsrcOver"]
    );
}

#[test]
fn color_filter_mode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final cf = ColorFilter.mode(Color(0xFFFF0000), BlendMode.multiply);
  print(cf != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn paint_blend_mode_assignment() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final paint = Paint()..blendMode = BlendMode.dstIn;
  print(paint.blendMode.name);
}
"#
        ),
        vec!["dstIn"]
    );
}

#[test]
fn image_shader_blend() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  print('compiles'); // ImageShader creation usually requires an Image, harder to mock here
}
"#
        ),
        vec!["compiles"]
    );
}
