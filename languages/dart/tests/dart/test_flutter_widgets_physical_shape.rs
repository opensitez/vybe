use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets PhysicalShape
// ═══════════════════════════════════════════════════════════

#[test]
fn physical_shape_creation() {
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
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF000000),
    child: const SizedBox(),
  );
  print(ps.clipper != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_shape_color() {
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
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF123456),
  );
  print(ps.color.value == 0xFF123456);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_shape_clip_behavior() {
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
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF000000),
    clipBehavior: Clip.hardEdge,
  );
  print(ps.clipBehavior == Clip.hardEdge);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn physical_shape_elevation() {
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
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF000000),
    elevation: 8.0,
  );
  print(ps.elevation);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn physical_shape_shadow_color() {
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
  final ps = PhysicalShape(
    clipper: MyPathClipper(),
    color: const Color(0xFF000000),
    shadowColor: const Color(0xFF222222),
  );
  print(ps.shadowColor.value == 0xFF222222);
}
"#
        ),
        vec!["true"]
    );
}
