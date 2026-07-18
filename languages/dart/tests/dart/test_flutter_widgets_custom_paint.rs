use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets CustomPaint
// ═══════════════════════════════════════════════════════════

#[test]
fn custom_paint_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
class MyPainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {}
  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
}
void main() {
  final cp = CustomPaint(painter: MyPainter());
  print(cp.painter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_foreground_painter() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
class MyPainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {}
  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
}
void main() {
  final cp = CustomPaint(foregroundPainter: MyPainter());
  print(cp.foregroundPainter != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = CustomPaint(size: Size(100, 200));
  print('${cp.size.width}:${cp.size.height}');
}
"#
        ),
        vec!["100.0:200.0"]
    );
}

#[test]
fn custom_paint_is_complex() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = CustomPaint(isComplex: true);
  print(cp.isComplex);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_will_change() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = CustomPaint(willChange: true);
  print(cp.willChange);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = CustomPaint(child: const Placeholder());
  print(cp.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = CustomPaint();
  print(cp is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_painter_semantics() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyPainter extends CustomPainter {
  @override
  void paint(Canvas canvas, Size size) {}
  @override
  bool shouldRepaint(covariant CustomPainter oldDelegate) => false;
  @override
  bool? hitTest(Offset position) => true;
}
void main() {
  final p = MyPainter();
  print(p.hitTest(Offset.zero));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn custom_paint_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final cp = CustomPaint();
  // Creates RenderCustomPaint
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
