// vybe-test: dart/dart_typed_data_int_lists/int32x4_operations_add
// origin: languages/dart/tests/dart/test_dart_typed_data_int_lists.rs

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
  final a = Int32x4(10, 20, 30, 40);
  final b = Int32x4(1, 2, 3, 4);
  final c = a + b;
  __p(c.y);
}

void main() {
  __vybeMain();
  __check('22');
}
