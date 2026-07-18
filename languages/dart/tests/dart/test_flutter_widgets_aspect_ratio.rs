use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets AspectRatio
// ═══════════════════════════════════════════════════════════

#[test]
fn aspect_ratio_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ar = AspectRatio(aspectRatio: 16 / 9);
  print(ar.aspectRatio.toStringAsFixed(2));
}
"#
        ),
        vec!["1.78"]
    );
}

#[test]
fn aspect_ratio_1_to_1() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ar = AspectRatio(aspectRatio: 1.0);
  print(ar.aspectRatio);
}
"#
        ),
        vec!["1.0"]
    );
}

#[test]
fn aspect_ratio_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ar = AspectRatio(aspectRatio: 2.0, child: const Placeholder());
  print(ar.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn aspect_ratio_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final ar = AspectRatio(aspectRatio: 1.0);
  print(ar is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn aspect_ratio_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final ar = AspectRatio(aspectRatio: 1.0);
  // Creates RenderAspectRatio
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}

#[test]
fn aspect_ratio_zero_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  try {
    AspectRatio(aspectRatio: 0.0);
  } catch (e) {
    print('error');
  }
}
"#
        ),
        vec!["error"]
    );
}

#[test]
fn aspect_ratio_negative_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  try {
    AspectRatio(aspectRatio: -1.0);
  } catch (e) {
    print('error');
  }
}
"#
        ),
        vec!["error"]
    );
}
