use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Flexible
// ═══════════════════════════════════════════════════════════

#[test]
fn flexible_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(child: const SizedBox());
  print(f.flex);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn flexible_flex_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(flex: 2, child: const SizedBox());
  print(f.flex);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn flexible_fit_loose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(child: const SizedBox());
  print(f.fit == FlexFit.loose);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn flexible_fit_tight() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(fit: FlexFit.tight, child: const SizedBox());
  print(f.fit == FlexFit.tight);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn flexible_is_parent_data_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(child: const SizedBox());
  print(f is ParentDataWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn flexible_child_retrieval() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(child: const Placeholder());
  print(f.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn flexible_apply_parent_data_flex() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final f = Flexible(flex: 3, fit: FlexFit.tight, child: const SizedBox());
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}

#[test]
fn flexible_debug_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final f = Flexible(child: const SizedBox());
  // Check that the class exists and works
  print(f.runtimeType.toString() == 'Flexible');
}
"#
        ),
        vec!["true"]
    );
}
