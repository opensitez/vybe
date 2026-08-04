// vybe-test: dart/dart_typed_data_float_lists/float64x2list_creation
// origin: languages/dart/tests/dart/test_dart_typed_data_float_lists.rs

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

import 'dart:typed_data';
void __vybeMain() {
  final l = Float64x2List(1);
  l[0] = Float64x2(9.9, 8.8);
  __p(l[0].y);
}

void main() {
  __vybeMain();
  __check('8.8');
}
