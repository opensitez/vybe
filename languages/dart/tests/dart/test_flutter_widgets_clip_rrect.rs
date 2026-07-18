use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ClipRRect
// ═══════════════════════════════════════════════════════════

#[test]
fn clip_rrect_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRRect(
    borderRadius: BorderRadius.circular(10.0),
    child: const SizedBox(),
  );
  print(cr.borderRadius != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rrect_border_radius() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRRect(
    borderRadius: BorderRadius.all(Radius.circular(20.0)),
  );
  print((cr.borderRadius as BorderRadius).topLeft.x);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn clip_rrect_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRRect(clipBehavior: Clip.antiAlias);
  print(cr.clipBehavior == Clip.antiAlias); // default
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rrect_clipper() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
class MyClipper extends CustomClipper<RRect> {
  @override
  RRect getClip(Size size) => RRect.fromLTRBR(0, 0, 50, 50, Radius.circular(5));
  @override
  bool shouldReclip(oldClipper) => false;
}
void main() {
  final cr = ClipRRect(clipper: MyClipper());
  print(cr.clipper != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_rrect_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cr = ClipRRect();
  print(cr is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
