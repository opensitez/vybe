use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: widgets SemanticsNode
// ═══════════════════════════════════════════════════════════

#[test]
fn semantics_node_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  print(sn != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn semantics_node_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.label = 'my_label';
  print(sn.label);
}
"#
        ),
        vec!["my_label"]
    );
}

#[test]
fn semantics_node_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.value = 'my_value';
  print(sn.value);
}
"#
        ),
        vec!["my_value"]
    );
}

#[test]
fn semantics_node_hint() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.hint = 'my_hint';
  print(sn.hint);
}
"#
        ),
        vec!["my_hint"]
    );
}

#[test]
fn semantics_node_increased_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.increasedValue = 'up';
  print(sn.increasedValue);
}
"#
        ),
        vec!["up"]
    );
}

#[test]
fn semantics_node_decreased_value() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.decreasedValue = 'down';
  print(sn.decreasedValue);
}
"#
        ),
        vec!["down"]
    );
}

#[test]
fn semantics_node_flags() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.hasCheckedState = true;
  sn.isChecked = true;
  print('${sn.hasCheckedState}:${sn.isChecked}');
}
"#
        ),
        vec!["true:true"]
    );
}

#[test]
fn semantics_node_rect() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
import 'dart:ui';
void main() {
  final sn = SemanticsNode();
  sn.rect = Rect.fromLTRB(0, 0, 100, 100);
  print(sn.rect.width);
}
"#
        ),
        vec!["100.0"]
    );
}

#[test]
fn semantics_node_tags() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final sn = SemanticsNode();
  sn.tags = {SemanticsTag('my_tag')};
  print(sn.tags!.first.name);
}
"#
        ),
        vec!["my_tag"]
    );
}

#[test]
fn semantics_configuration() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final config = SemanticsConfiguration();
  config.label = 'config_label';
  print(config.label);
}
"#
        ),
        vec!["config_label"]
    );
}

#[test]
fn semantics_properties() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final props = SemanticsProperties(label: 'props_label');
  print(props.label);
}
"#
        ),
        vec!["props_label"]
    );
}

#[test]
fn custom_semantics_action() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/semantics.dart';
void main() {
  final action = CustomSemanticsAction.overridableAction(hint: 'my_action', action: SemanticsAction.tap);
  print(action.hint);
}
"#
        ),
        vec!["my_action"]
    );
}
