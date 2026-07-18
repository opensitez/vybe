use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Baseline
// ═══════════════════════════════════════════════════════════

#[test]
fn baseline_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final b = Baseline(
    baseline: 20.0,
    baselineType: TextBaseline.alphabetic,
    child: const SizedBox(),
  );
  print(b.baseline);
}
"#
        ),
        vec!["20.0"]
    );
}

#[test]
fn baseline_type_alphabetic() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final b = Baseline(
    baseline: 10.0,
    baselineType: TextBaseline.alphabetic,
  );
  print(b.baselineType == TextBaseline.alphabetic);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn baseline_type_ideographic() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final b = Baseline(
    baseline: 15.0,
    baselineType: TextBaseline.ideographic,
  );
  print(b.baselineType == TextBaseline.ideographic);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn baseline_child() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final b = Baseline(
    baseline: 0.0,
    baselineType: TextBaseline.alphabetic,
    child: const Placeholder(),
  );
  print(b.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn baseline_is_single_child_render_object_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final b = Baseline(
    baseline: 10.0,
    baselineType: TextBaseline.alphabetic,
  );
  print(b is SingleChildRenderObjectWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn text_baseline_values() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  print(TextBaseline.alphabetic.name);
  print(TextBaseline.ideographic.name);
}
"#
        ),
        vec!["alphabetic\nideographic"]
    );
}
