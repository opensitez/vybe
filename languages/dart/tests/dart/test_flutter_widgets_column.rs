use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Column
// ═══════════════════════════════════════════════════════════

#[test]
fn column_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(children: [const SizedBox(), const SizedBox()]);
  print(c.children.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn column_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column();
  print(c.direction == Axis.vertical);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_main_axis_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(mainAxisAlignment: MainAxisAlignment.end);
  print(c.mainAxisAlignment == MainAxisAlignment.end);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_main_axis_size() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(mainAxisSize: MainAxisSize.max);
  print(c.mainAxisSize == MainAxisSize.max);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_cross_axis_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(crossAxisAlignment: CrossAxisAlignment.baseline);
  print(c.crossAxisAlignment == CrossAxisAlignment.baseline);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_text_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(textDirection: TextDirection.ltr);
  print(c.textDirection == TextDirection.ltr);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_vertical_direction() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(verticalDirection: VerticalDirection.down);
  print(c.verticalDirection == VerticalDirection.down);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_text_baseline() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(textBaseline: TextBaseline.ideographic);
  print(c.textBaseline == TextBaseline.ideographic);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_clip_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column(clipBehavior: Clip.none);
  print(c.clipBehavior == Clip.none);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn column_is_flex() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final c = Column();
  print(c is Flex);
}
"#
        ),
        vec!["true"]
    );
}
