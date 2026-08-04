// vybe-test: dart/flutter_foundation_global_key/global_key_inequality
// origin: languages/dart/tests/dart/test_flutter_foundation_global_key.rs

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

import 'package:flutter/widgets.dart';
void __vybeMain() {
  final k1 = GlobalKey();
  final k2 = GlobalKey();
  __p(k1 == k2);
}

void main() {
  __vybeMain();
  __check('false');
}
