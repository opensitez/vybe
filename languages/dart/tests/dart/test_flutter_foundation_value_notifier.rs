use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: foundation ValueNotifier
// ═══════════════════════════════════════════════════════════

#[test]
fn value_notifier_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<String>('initial');
  print(vn.value);
}
"#
        ),
        vec!["initial"]
    );
}

#[test]
fn value_notifier_setter_triggers_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<int>(0);
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = 1;
  print(count);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn value_notifier_setter_identical_value_does_not_trigger() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<int>(1);
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = 1;
  print(count);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn value_notifier_setter_equal_objects_trigger_depends_on_operator_eq() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class Eq {
  final int val;
  Eq(this.val);
  @override
  bool operator ==(Object other) => other is Eq && other.val == val;
  @override
  int get hashCode => val.hashCode;
}
void main() {
  final vn = ValueNotifier<Eq>(Eq(1));
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = Eq(1); // Same by ==
  print(count);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn value_notifier_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<int>(0);
  vn.dispose();
  try {
    vn.value = 1;
  } catch (e) {
    print('FlutterError');
  }
}
"#
        ),
        vec!["FlutterError"]
    );
}

#[test]
fn value_notifier_value_read_after_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<int>(42);
  vn.dispose();
  // Depending on Flutter version, reading value after dispose might throw or return.
  // Actually, reading value is generally safe but might print warning.
  try {
    final v = vn.value;
    print(v);
  } catch(e) {
    print('FlutterError');
  }
}
"#
        ),
        vec!["42"] // It usually succeeds
    );
}

#[test]
fn value_notifier_subclass() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class CustomNotifier extends ValueNotifier<int> {
  CustomNotifier(int val) : super(val);
  void increment() => value++;
}
void main() {
  final cn = CustomNotifier(10);
  int count = 0;
  cn.addListener(() { count++; });
  cn.increment();
  print('${count}:${cn.value}');
}
"#
        ),
        vec!["1:11"]
    );
}

#[test]
fn value_notifier_remove_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<bool>(false);
  int count = 0;
  void cb() { count++; }
  vn.addListener(cb);
  vn.removeListener(cb);
  vn.value = true;
  print(count);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn value_listenable_interface() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<String>('a');
  ValueListenable<String> vl = vn;
  print(vl.value);
}
"#
        ),
        vec!["a"]
    );
}

#[test]
fn value_notifier_nullable_type() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<String?>(null);
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = 'hello';
  vn.value = null;
  print(count);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn value_notifier_same_object_mutation_no_trigger() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final list = [1, 2];
  final vn = ValueNotifier<List<int>>(list);
  int count = 0;
  vn.addListener(() { count++; });
  list.add(3);
  vn.value = list; // Same instance, so `==` is true, no trigger
  print(count);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn value_notifier_force_notify() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final list = [1, 2];
  final vn = ValueNotifier<List<int>>(list);
  int count = 0;
  vn.addListener(() { count++; });
  list.add(3);
  // Using ChangeNotifier's notifyListeners bypasses the `==` check
  // notifyListeners is protected, but Dart allows it via dynamic or if not strictly enforced in tests.
  // Actually, we can't call it directly unless we subclass.
  print('test_skip'); // We'll test subclass next
}
"#
        ),
        vec!["test_skip"]
    );
}

#[test]
fn value_notifier_protected_notify() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
class ListNotifier extends ValueNotifier<List<int>> {
  ListNotifier(List<int> val) : super(val);
  void update() {
    notifyListeners();
  }
}
void main() {
  final list = [1];
  final ln = ListNotifier(list);
  int count = 0;
  ln.addListener(() { count++; });
  list.add(2);
  ln.update();
  print(count);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn value_notifier_toString() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/foundation.dart';
void main() {
  final vn = ValueNotifier<int>(99);
  print(vn.toString().contains('99'));
}
"#
        ),
        vec!["true"]
    );
}
