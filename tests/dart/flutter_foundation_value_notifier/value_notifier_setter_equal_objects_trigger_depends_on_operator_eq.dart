// vybe-test: dart/flutter_foundation_value_notifier/value_notifier_setter_equal_objects_trigger_depends_on_operator_eq
// origin: languages/dart/tests/dart/test_flutter_foundation_value_notifier.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'package:flutter/foundation.dart';
class Eq {
  final int val;
  Eq(this.val);
  @override
  bool operator ==(Object other) => other is Eq && other.val == val;
  @override
  int get hashCode => val.hashCode;
}
void __vybeMain() {
  final vn = ValueNotifier<Eq>(Eq(1));
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = Eq(1); // Same by ==
  __p(count);
}

void main() {
  __vybeMain();
  __check('0');
}
