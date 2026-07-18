use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Transform
// ═══════════════════════════════════════════════════════════

#[test]
fn transform_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform(transform: Matrix4.identity());
  print(t.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_rotate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:math';
void main() {
  final t = Transform.rotate(angle: pi / 2);
  print(t.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_translate() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:math';
void main() {
  final t = Transform.translate(offset: const Offset(10, 20));
  print(t.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_scale() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform.scale(scale: 2.0);
  print(t.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_scale_xy() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform.scale(scaleX: 2.0, scaleY: 3.0);
  print(t.transform != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_origin() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform(
    transform: Matrix4.identity(),
    origin: const Offset(50, 50),
  );
  print(t.origin == const Offset(50, 50));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform(
    transform: Matrix4.identity(),
    alignment: Alignment.center,
  );
  print(t.alignment == Alignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transform_transform_hit_tests() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform(
    transform: Matrix4.identity(),
    transformHitTests: false,
  );
  print(t.transformHitTests);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn transform_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final t = Transform(transform: Matrix4.identity());
  print(t is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}
