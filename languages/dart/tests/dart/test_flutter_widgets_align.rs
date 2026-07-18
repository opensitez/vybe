use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Align
// ═══════════════════════════════════════════════════════════

#[test]
fn align_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(child: const SizedBox());
  print(a.alignment == Alignment.center); // default is center
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn align_alignment_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(alignment: Alignment.topLeft, child: const SizedBox());
  print(a.alignment == Alignment.topLeft);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn align_width_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(widthFactor: 2.0, child: const SizedBox());
  print(a.widthFactor);
}
"#
        ),
        vec!["2.0"]
    );
}

#[test]
fn align_height_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(heightFactor: 0.5, child: const SizedBox());
  print(a.heightFactor);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn align_is_single_child_render_object_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align();
  print(a is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn align_fractional_offset() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(alignment: const FractionalOffset(0.2, 0.8));
  print(a.alignment == const FractionalOffset(0.2, 0.8));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn align_directional_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = Align(alignment: AlignmentDirectional.bottomStart);
  print(a.alignment == AlignmentDirectional.bottomStart);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_align_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = AnimatedAlign(
    alignment: Alignment.topRight,
    duration: const Duration(seconds: 1),
    child: const SizedBox(),
  );
  print(a.duration.inSeconds);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn animated_align_curve() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = AnimatedAlign(
    alignment: Alignment.center,
    duration: const Duration(seconds: 1),
    curve: Curves.easeIn,
  );
  print(a.curve == Curves.easeIn);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn animated_align_is_implicit_animation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final a = AnimatedAlign(
    alignment: Alignment.center,
    duration: const Duration(seconds: 1),
  );
  print(a is ImplicitlyAnimatedWidget);
}
"#
        ),
        vec!["true"]
    );
}
