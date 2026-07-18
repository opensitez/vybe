use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Row
// ═══════════════════════════════════════════════════════════

#[test]
fn row_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(children: [const SizedBox()]);
  print(r.children.length);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn row_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row();
  print(r.direction == Axis.horizontal);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_main_axis_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(mainAxisAlignment: MainAxisAlignment.spaceBetween);
  print(r.mainAxisAlignment == MainAxisAlignment.spaceBetween);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_main_axis_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(mainAxisSize: MainAxisSize.min);
  print(r.mainAxisSize == MainAxisSize.min);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_cross_axis_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(crossAxisAlignment: CrossAxisAlignment.stretch);
  print(r.crossAxisAlignment == CrossAxisAlignment.stretch);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(textDirection: TextDirection.rtl);
  print(r.textDirection == TextDirection.rtl);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_vertical_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(verticalDirection: VerticalDirection.up);
  print(r.verticalDirection == VerticalDirection.up);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_text_baseline() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(textBaseline: TextBaseline.alphabetic);
  print(r.textBaseline == TextBaseline.alphabetic);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row(clipBehavior: Clip.hardEdge);
  print(r.clipBehavior == Clip.hardEdge);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn row_is_flex() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final r = Row();
  print(r is Flex);
}
"#
        ),
        vec!["true"]
    );
}
