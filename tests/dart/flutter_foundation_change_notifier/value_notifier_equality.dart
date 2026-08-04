// vybe-test: dart/flutter_foundation_change_notifier/value_notifier_equality
// origin: languages/dart/tests/dart/test_flutter_foundation_change_notifier.rs

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
  final v1 = ValueNotifier<int>(1);
  final v2 = ValueNotifier<int>(1);
  __p(v1 == v2);
}

void main() {
  __vybeMain();
  __check('false');
}
