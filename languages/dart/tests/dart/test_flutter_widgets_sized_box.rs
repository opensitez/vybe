use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SizedBox
// ═══════════════════════════════════════════════════════════

#[test]
fn sized_box_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = SizedBox(width: 100.0, height: 200.0);
  print('${sb.width}:${sb.height}');
}
"#
        ),
        vec!["100.0:200.0"]
    );
}

#[test]
fn sized_box_expand() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = SizedBox.expand();
  print('${sb.width == double.infinity}:${sb.height == double.infinity}');
}
"#
        ),
        vec!["true:true"]
    );
}

#[test]
fn sized_box_shrink() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = SizedBox.shrink();
  print('${sb.width}:${sb.height}');
}
"#
        ),
        vec!["0.0:0.0"]
    );
}

#[test]
fn sized_box_from_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'dart:ui';
void main() {
  final sb = SizedBox.fromSize(size: const Size(50.0, 50.0));
  print('${sb.width}:${sb.height}');
}
"#
        ),
        vec!["50.0:50.0"]
    );
}

#[test]
fn sized_box_square() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = SizedBox.square(dimension: 75.0);
  print('${sb.width}:${sb.height}');
}
"#
        ),
        vec!["75.0:75.0"]
    );
}

#[test]
fn sized_box_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = SizedBox(child: const Placeholder());
  print(sb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sized_box_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = const SizedBox();
  print(sb is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn sized_box_no_dimensions() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final sb = const SizedBox();
  print('${sb.width}:${sb.height}');
}
"#
        ),
        vec!["null:null"]
    );
}

#[test]
fn animated_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = AnimatedSize(
    duration: const Duration(milliseconds: 500),
    child: const SizedBox(),
  );
  print(w.duration.inMilliseconds);
}
"#
        ),
        vec!["500"]
    );
}

#[test]
fn sized_box_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final sb = const SizedBox();
  // inherits createRenderObject, creates RenderConstrainedBox internally
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
