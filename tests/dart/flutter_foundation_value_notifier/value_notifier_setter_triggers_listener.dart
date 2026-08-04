// vybe-test: dart/flutter_foundation_value_notifier/value_notifier_setter_triggers_listener
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
void __vybeMain() {
  final vn = ValueNotifier<int>(0);
  int count = 0;
  vn.addListener(() { count++; });
  vn.value = 1;
  __p(count);
}

void main() {
  __vybeMain();
  __check('1');
}
