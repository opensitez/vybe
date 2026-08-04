// vybe-test: dart/flutter_foundation_key_local_key/value_key_hashcode
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
void __vybeMain() {
  final k1 = Key('test');
  final k2 = Key('test');
  __p(k1.hashCode == k2.hashCode);
}

void main() {
  __vybeMain();
  __check('true');
}
