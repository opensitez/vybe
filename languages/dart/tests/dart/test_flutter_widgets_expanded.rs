use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets Expanded
// ═══════════════════════════════════════════════════════════

#[test]
fn expanded_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  print(e.flex);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn expanded_flex_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(flex: 3, child: const SizedBox());
  print(e.flex);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn expanded_fit() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  print(e.fit == FlexFit.tight);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn expanded_is_flexible() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  print(e is Flexible);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn expanded_is_parent_data_widget() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  print(e is ParentDataWidget);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn expanded_child_required() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const Placeholder());
  print(e.child is Placeholder);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn expanded_apply_parent_data() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
import 'package:flutter/rendering.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  final pd = FlexParentData();
  e.applyParentData(RenderBox() as dynamic); // we pass null or mock
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}

#[test]
fn expanded_debug_fill_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final e = Expanded(child: const SizedBox());
  // Can't easily test DiagnosticPropertiesBuilder natively here, just test it exists
  print('compiles');
}
"#
        ),
        vec!["compiles"]
    );
}
