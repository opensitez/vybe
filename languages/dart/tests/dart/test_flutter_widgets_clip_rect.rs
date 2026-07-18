use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ClipRect
// ═══════════════════════════════════════════════════════════

#[test]
fn clip_rect_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRect(child: const SizedBox());
  print(cr.clipBehavior == Clip.hardEdge);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rect_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRect(clipBehavior: Clip.antiAlias, child: const SizedBox());
  print(cr.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rect_clipper() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
class MyClipper extends CustomClipper<Rect> {
  @override
  Rect getClip(Size size) => Rect.fromLTWH(0, 0, 50, 50);
  @override
  bool shouldReclip(oldClipper) => false;
}
void main() {
  final cr = ClipRect(clipper: MyClipper());
  print(cr.clipper != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rect_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRect();
  print(cr is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_behavior_values() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  print(Clip.none.name);
  print(Clip.hardEdge.name);
  print(Clip.antiAlias.name);
  print(Clip.antiAliasWithSaveLayer.name);
}
"#
        ),
        vec!["none\nhardEdge\nantiAlias\nantiAliasWithSaveLayer"]
    );
}
