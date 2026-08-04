// vybe-test: dart/flutter_foundation_key_local_key/object_key_identity_inequality
// origin: languages/dart/tests/dart/test_flutter_foundation_key_local_key.rs

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
class MyObj {
  @override
  bool operator ==(Object other) => true; // Always equal
}
void __vybeMain() {
  final obj1 = MyObj();
  final obj2 = MyObj();
  // ObjectKey uses identical(), not ==
  final k1 = ObjectKey(obj1);
  final k2 = ObjectKey(obj2);
  __p(k1 == k2);
}

void main() {
  __vybeMain();
  __check('false');
}
