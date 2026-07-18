use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: foundation Key & LocalKey
// ═══════════════════════════════════════════════════════════

#[test]
fn key_creation_value_key() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k = Key('my_key');
  print(k is ValueKey<String>);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_key_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = Key('test');
  final k2 = Key('test');
  print(k1 == k2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_key_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = Key('test1');
  final k2 = Key('test2');
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn value_key_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = Key('test');
  final k2 = Key('test');
  print(k1.hashCode == k2.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_key_integer() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = ValueKey<int>(42);
  final k2 = ValueKey<int>(42);
  print(k1 == k2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_key_type_mismatch() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = ValueKey<String>('42');
  final k2 = ValueKey<int>(42);
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn object_key_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyObj {}
void main() {
  final obj = MyObj();
  final k1 = ObjectKey(obj);
  final k2 = ObjectKey(obj);
  print(k1 == k2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn object_key_identity_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyObj {
  @override
  bool operator ==(Object other) => true; // Always equal
}
void main() {
  final obj1 = MyObj();
  final obj2 = MyObj();
  // ObjectKey uses identical(), not ==
  final k1 = ObjectKey(obj1);
  final k2 = ObjectKey(obj2);
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn object_key_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyObj {}
void main() {
  final obj = MyObj();
  final k1 = ObjectKey(obj);
  final k2 = ObjectKey(obj);
  print(k1.hashCode == k2.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn local_key_subclass() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class MyLocalKey extends LocalKey {
  const MyLocalKey();
}
void main() {
  final k = const MyLocalKey();
  print(k is Key);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn unique_key_inequality() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = UniqueKey();
  final k2 = UniqueKey();
  print(k1 == k2);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn unique_key_equality_to_self() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k1 = UniqueKey();
  print(k1 == k1);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn key_empty_string() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k = Key('');
  print(k.value == '');
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn object_key_null() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k = ObjectKey(null);
  print(k.value == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn value_key_toString() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final k = Key('abc');
  print(k.toString().contains('abc'));
}
"#
        ),
        vec!["true"]
    );
}
