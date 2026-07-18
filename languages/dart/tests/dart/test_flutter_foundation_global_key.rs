use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: foundation GlobalKey
// ═══════════════════════════════════════════════════════════

#[test]
fn global_key_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey();
  print(k is Key);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_key_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k1 = GlobalKey();
  final k2 = GlobalKey();
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn global_key_equality_to_self() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey();
  print(k == k);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_key_with_label() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey(debugLabel: 'my_label');
  print(k.toString().contains('my_label'));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn labeled_global_key_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k1 = GlobalKey(debugLabel: 'A');
  final k2 = GlobalKey(debugLabel: 'A');
  // Labels are just for debugging, they don't make them equal
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn global_key_typed() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey<State<StatefulWidget>>();
  print(k is GlobalKey);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_object_key_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyObj {}
void main() {
  final obj = MyObj();
  final k = GlobalObjectKey(obj);
  print(k.value == obj);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_object_key_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyObj {}
void main() {
  final obj = MyObj();
  final k1 = GlobalObjectKey(obj);
  final k2 = GlobalObjectKey(obj);
  print(k1 == k2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_object_key_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyObj {}
void main() {
  final k1 = GlobalObjectKey(MyObj());
  final k2 = GlobalObjectKey(MyObj());
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn global_object_key_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
class MyObj {}
void main() {
  final obj = MyObj();
  final k1 = GlobalObjectKey(obj);
  final k2 = GlobalObjectKey(obj);
  print(k1.hashCode == k2.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_key_current_context_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey();
  // Not attached to any tree
  print(k.currentContext == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_key_current_widget_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey();
  print(k.currentWidget == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn global_key_current_state_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/widgets.dart';
void main() {
  final k = GlobalKey<State<StatefulWidget>>();
  print(k.currentState == null);
}
"#
        ),
        vec!["true"]
    );
}
