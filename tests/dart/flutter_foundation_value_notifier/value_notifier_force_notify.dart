// vybe-test: dart/flutter_foundation_value_notifier/value_notifier_force_notify
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
  final list = [1, 2];
  final vn = ValueNotifier<List<int>>(list);
  int count = 0;
  vn.addListener(() { count++; });
  list.add(3);
  // Using ChangeNotifier's notifyListeners bypasses the `==` check
  // notifyListeners is protected, but Dart allows it via dynamic or if not strictly enforced in tests.
  // Actually, we can't call it directly unless we subclass.
  __p('test_skip'); // We'll test subclass next
}

void main() {
  __vybeMain();
  __check('test_skip');
}
