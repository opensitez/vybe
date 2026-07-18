use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ClipOval
// ═══════════════════════════════════════════════════════════

#[test]
fn clip_oval_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final co = ClipOval(child: const SizedBox());
  print(co.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_oval_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final co = ClipOval(clipBehavior: Clip.hardEdge, child: const SizedBox());
  print(co.clipBehavior == Clip.hardEdge);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_oval_clipper() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
class MyOvalClipper extends CustomClipper<Rect> {
  @override
  Rect getClip(Size size) => Rect.fromLTWH(0,0,50,50);
  @override
  bool shouldReclip(oldClipper) => false;
}
void main() {
  final co = ClipOval(clipper: MyOvalClipper());
  print(co.clipper != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_oval_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final co = ClipOval();
  print(co is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
