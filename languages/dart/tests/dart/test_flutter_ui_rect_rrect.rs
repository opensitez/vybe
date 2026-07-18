use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: dart:ui Rect & RRect
// ═══════════════════════════════════════════════════════════

#[test]
fn rect_from_ltrb() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(10.0, 20.0, 30.0, 40.0);
  print('${r.left}:${r.top}:${r.right}:${r.bottom}');
}
"#
        ),
        vec!["10.0:20.0:30.0:40.0"]
    );
}

#[test]
fn rect_from_ltwh() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTWH(10.0, 20.0, 20.0, 20.0);
  print('${r.right}:${r.bottom}');
}
"#
        ),
        vec!["30.0:40.0"]
    );
}

#[test]
fn rect_from_circle() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final center = Offset(10.0, 10.0);
  final r = Rect.fromCircle(center: center, radius: 5.0);
  print('${r.left}:${r.right}:${r.top}:${r.bottom}');
}
"#
        ),
        vec!["5.0:15.0:5.0:15.0"]
    );
}

#[test]
fn rect_from_points() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final p1 = Offset(20.0, 30.0);
  final p2 = Offset(10.0, 10.0); // Out of order coordinates
  final r = Rect.fromPoints(p1, p2);
  print('${r.left}:${r.right}:${r.top}:${r.bottom}');
}
"#
        ),
        vec!["10.0:20.0:10.0:30.0"]
    );
}

#[test]
fn rect_properties_width_height() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(0.0, 0.0, 100.0, 50.0);
  print('${r.width}:${r.height}');
}
"#
        ),
        vec!["100.0:50.0"]
    );
}

#[test]
fn rect_center() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  print('${r.center.dx}:${r.center.dy}');
}
"#
        ),
        vec!["5.0:5.0"]
    );
}

#[test]
fn rect_is_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r1 = Rect.fromLTRB(0.0, 0.0, 0.0, 0.0);
  final r2 = Rect.fromLTRB(10.0, 0.0, 0.0, 0.0); // right < left
  print(r1.isEmpty);
  print(r2.isEmpty);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn rect_contains() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  print(r.contains(Offset(5.0, 5.0)));
  print(r.contains(Offset(15.0, 5.0)));
}
"#
        ),
        vec!["true\nfalse"]
    );
}

#[test]
fn rect_intersect() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r1 = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final r2 = Rect.fromLTRB(5.0, 5.0, 15.0, 15.0);
  final inter = r1.intersect(r2);
  print('${inter.left}:${inter.top}:${inter.right}:${inter.bottom}');
}
"#
        ),
        vec!["5.0:5.0:10.0:10.0"]
    );
}

#[test]
fn rect_expand_to_include() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r1 = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final r2 = Rect.fromLTRB(5.0, 5.0, 15.0, 15.0);
  final expanded = r1.expandToInclude(r2);
  print('${expanded.left}:${expanded.right}:${expanded.bottom}');
}
"#
        ),
        vec!["0.0:15.0:15.0"]
    );
}

#[test]
fn rect_overlaps() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r1 = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final r2 = Rect.fromLTRB(10.0, 10.0, 20.0, 20.0); // Touches edge
  // Overlaps strictly greater than
  print(r1.overlaps(r2));
}
"#
        ),
        // Overlaps requires strict intersection interior
        vec!["false"]
    );
}

#[test]
fn rect_inflate() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(10.0, 10.0, 20.0, 20.0);
  final inf = r.inflate(5.0);
  print('${inf.left}:${inf.right}');
}
"#
        ),
        vec!["5.0:25.0"]
    );
}

#[test]
fn rect_deflate() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final r = Rect.fromLTRB(10.0, 10.0, 20.0, 20.0);
  final def = r.deflate(2.0);
  print('${def.left}:${def.right}');
}
"#
        ),
        vec!["12.0:18.0"]
    );
}

#[test]
fn rrect_from_rect_and_radius() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final rect = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final rrect = RRect.fromRectAndRadius(rect, Radius.circular(2.0));
  print(rrect.tlRadiusX);
}
"#
        ),
        vec!["2.0"]
    );
}

#[test]
fn rrect_from_rect_and_corners() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final rect = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final rrect = RRect.fromRectAndCorners(rect, topLeft: Radius.circular(5.0));
  print('${rrect.tlRadiusX}:${rrect.trRadiusX}');
}
"#
        ),
        vec!["5.0:0.0"]
    );
}

#[test]
fn rrect_contains() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:ui';
void main() {
  final rect = Rect.fromLTRB(0.0, 0.0, 10.0, 10.0);
  final rrect = RRect.fromRectAndRadius(rect, Radius.circular(5.0)); // A circle
  print(rrect.contains(Offset(5.0, 5.0))); // Center
  print(rrect.contains(Offset(0.0, 0.0))); // Corner (outside circle)
}
"#
        ),
        vec!["true\nfalse"]
    );
}
