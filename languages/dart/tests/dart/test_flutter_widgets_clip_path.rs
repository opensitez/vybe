use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets ClipPath
// ═══════════════════════════════════════════════════════════

#[test]
fn clip_path_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = ClipPath(child: const SizedBox());
  print(cp.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_path_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = ClipPath(clipBehavior: Clip.none, child: const SizedBox());
  print(cp.clipBehavior == Clip.none);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_path_clipper() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
class MyPathClipper extends CustomClipper<Path> {
  @override
  Path getClip(Size size) => Path()..addRect(Rect.fromLTWH(0,0,50,50));
  @override
  bool shouldReclip(oldClipper) => false;
}
void main() {
  final cp = ClipPath(clipper: MyPathClipper());
  print(cp.clipper != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn clip_path_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final cp = ClipPath();
  print(cp is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
