use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets FractionallySizedBox
// ═══════════════════════════════════════════════════════════

#[test]
fn fractionally_sized_box_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox();
  print(f != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractionally_sized_box_width_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox(widthFactor: 0.5);
  print(f.widthFactor);
}
"#
        ),
        vec!["0.5"]
    );
}

#[test]
fn fractionally_sized_box_height_factor() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox(heightFactor: 0.8);
  print(f.heightFactor);
}
"#
        ),
        vec!["0.8"]
    );
}

#[test]
fn fractionally_sized_box_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox(alignment: Alignment.topRight);
  print(f.alignment == Alignment.topRight);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractionally_sized_box_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox(child: const SizedBox());
  print(f.child is SizedBox);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractionally_sized_box_is_single_child_render_object() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox();
  print(f is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractionally_sized_box_default_alignment() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox();
  print(f.alignment == Alignment.center);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn fractionally_sized_box_factors_null_by_default() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox();
  print('${f.widthFactor}:${f.heightFactor}');
}
"#
        ),
        vec!["null:null"]
    );
}

#[test]
fn fractionally_sized_box_over_one() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = FractionallySizedBox(widthFactor: 1.5, heightFactor: 2.0);
  print('${f.widthFactor}:${f.heightFactor}');
}
"#
        ),
        vec!["1.5:2.0"]
    );
}
