use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Wrap
// ═══════════════════════════════════════════════════════════

#[test]
fn wrap_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(children: [const SizedBox()]);
  print(w.children.length);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn wrap_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap();
  print(w.direction == Axis.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_vertical_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(direction: Axis.vertical);
  print(w.direction == Axis.vertical);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(alignment: WrapAlignment.center);
  print(w.alignment == WrapAlignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_spacing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(spacing: 8.0);
  print(w.spacing);
}
"#
        ),
        vec!["8.0"]
    );
}

#[test]
fn wrap_run_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(runAlignment: WrapAlignment.end);
  print(w.runAlignment == WrapAlignment.end);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_run_spacing() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(runSpacing: 4.0);
  print(w.runSpacing);
}
"#
        ),
        vec!["4.0"]
    );
}

#[test]
fn wrap_cross_axis_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(crossAxisAlignment: WrapCrossAlignment.center);
  print(w.crossAxisAlignment == WrapCrossAlignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(textDirection: TextDirection.rtl);
  print(w.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_vertical_direction_property() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(verticalDirection: VerticalDirection.up);
  print(w.verticalDirection == VerticalDirection.up);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn wrap_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final w = Wrap(clipBehavior: Clip.antiAlias);
  print(w.clipBehavior == Clip.antiAlias);
}
"#
        ),
        vec!["true"]
    );
}
