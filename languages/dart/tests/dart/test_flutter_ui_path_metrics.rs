use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: dart:ui Path & PathMetrics
// ═══════════════════════════════════════════════════════════

#[test]
fn path_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  print(path != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_move_to_line_to() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.lineTo(10, 10);
  final bounds = path.getBounds();
  print('${bounds.width}:${bounds.height}');
}
"#
        ),
        vec!["10.0:10.0"]
    );
}

#[test]
fn path_add_rect() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  print(path.contains(Offset(5, 5)));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_add_oval() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addOval(Rect.fromLTRB(0, 0, 10, 10));
  print(path.contains(Offset(5, 5)));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_add_polygon() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addPolygon([Offset(0, 0), Offset(10, 0), Offset(5, 10)], true);
  print(path.contains(Offset(5, 5)));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_close() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.lineTo(10, 0);
  path.lineTo(10, 10);
  path.close();
  print(path.contains(Offset(5, 5))); // A triangle
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_reset() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  path.reset();
  print(path.getBounds().isEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_shift() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  final shifted = path.shift(Offset(5, 5));
  print('${shifted.getBounds().left}:${shifted.getBounds().top}');
}
"#
        ),
        vec!["5.0:5.0"]
    );
}

#[test]
fn path_transform() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
import 'dart:typed_data';
void main() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  // Identity matrix is 16 elements
  final matrix = Float64List(16);
  matrix[0] = 2.0; // scale X
  matrix[5] = 2.0; // scale Y
  matrix[10] = 1.0;
  matrix[15] = 1.0;
  
  final transformed = path.transform(matrix);
  print(transformed.getBounds().width);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn path_metrics_length() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.lineTo(10, 0);
  final metrics = path.computeMetrics().toList();
  print(metrics.length == 1);
  print(metrics[0].length);
}
"#
        ),
        vec!["true\n10.0"]
    );
}

#[test]
fn path_metrics_is_closed() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.addRect(Rect.fromLTRB(0, 0, 10, 10));
  final metrics = path.computeMetrics().toList();
  print(metrics[0].isClosed);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_metrics_extract_path() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.lineTo(10, 0);
  final metrics = path.computeMetrics().toList();
  final subPath = metrics[0].extractPath(0, 5);
  print(subPath.getBounds().width);
}
"#
        ),
        vec!["5.0"]
    );
}

#[test]
fn path_metrics_tangent_for_offset() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.lineTo(10, 0);
  final metrics = path.computeMetrics().toList();
  final tangent = metrics[0].getTangentForOffset(5);
  print('${tangent!.position.dx}:${tangent.vector.dx}');
}
"#
        ),
        vec!["5.0:1.0"]
    );
}

#[test]
fn path_fill_type() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.fillType = PathFillType.evenOdd;
  print(path.fillType == PathFillType.evenOdd);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_combine_union() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final p1 = Path()..addRect(Rect.fromLTRB(0, 0, 10, 10));
  final p2 = Path()..addRect(Rect.fromLTRB(5, 5, 15, 15));
  final combined = Path.combine(PathOperation.union, p1, p2);
  print('${combined.getBounds().width}:${combined.getBounds().height}');
}
"#
        ),
        vec!["15.0:15.0"]
    );
}

#[test]
fn path_combine_intersect() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final p1 = Path()..addRect(Rect.fromLTRB(0, 0, 10, 10));
  final p2 = Path()..addRect(Rect.fromLTRB(5, 5, 15, 15));
  final combined = Path.combine(PathOperation.intersect, p1, p2);
  print('${combined.getBounds().width}:${combined.getBounds().height}');
}
"#
        ),
        vec!["5.0:5.0"]
    );
}

#[test]
fn path_add_arc() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  // Math.pi is 3.14...
  path.addArc(Rect.fromLTRB(0, 0, 10, 10), 0, 3.141592653589793);
  print(path.getBounds().isEmpty == false);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_conic_to() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.conicTo(5, 10, 10, 0, 1.0);
  print(path.getBounds().height > 0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn path_bezier_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(0, 0);
  path.cubicTo(2, 5, 8, 5, 10, 0);
  print(path.getBounds().width);
}
"#
        ),
        vec!["10.0"]
    );
}

#[test]
fn path_relative_line_to() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final path = Path();
  path.moveTo(10, 10);
  path.relativeLineTo(5, 5); // Goes to 15, 15
  print(path.getBounds().right);
}
"#
        ),
        vec!["15.0"]
    );
}
