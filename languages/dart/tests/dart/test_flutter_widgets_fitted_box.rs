use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets FittedBox
// ═══════════════════════════════════════════════════════════

#[test]
fn fitted_box_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox();
  print(fb != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_fit() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox(fit: BoxFit.contain);
  print(fb.fit == BoxFit.contain);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox(alignment: Alignment.bottomRight);
  print(fb.alignment == Alignment.bottomRight);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox(clipBehavior: Clip.hardEdge);
  print(fb.clipBehavior == Clip.hardEdge);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox(child: const Placeholder());
  print(fb.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox();
  print(fb is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fitted_box_default_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final fb = FittedBox();
  print('${fb.fit.name}:${fb.alignment == Alignment.center}');
}
"#
        ),
        vec!["contain:true"]
    );
}

#[test]
fn box_fit_values() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  print(BoxFit.fill.name);
  print(BoxFit.cover.name);
  print(BoxFit.scaleDown.name);
}
"#
        ),
        vec!["fill\ncover\nscaleDown"]
    );
}

#[test]
fn fitted_sizes() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/painting.dart';
import 'dart:ui';
void main() {
  final fs = FittedSizes(Size(10, 10), Size(20, 20));
  print('${fs.source.width}:${fs.destination.width}');
}
"#
        ),
        vec!["10.0:20.0"]
    );
}

#[test]
fn apply_box_fit() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/painting.dart';
import 'dart:ui';
void main() {
  final fs = applyBoxFit(BoxFit.contain, Size(100, 100), Size(50, 50));
  print('${fs.source.width}:${fs.destination.width}');
}
"#
        ),
        vec!["100.0:50.0"]
    );
}
