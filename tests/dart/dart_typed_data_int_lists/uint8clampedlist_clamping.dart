// vybe-test: dart/dart_typed_data_int_lists/uint8clampedlist_clamping
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
  final l = Uint8ClampedList(2);
  l[0] = 300; // Clamped to 255
  l[1] = -50; // Clamped to 0
  __p('${l[0]}:${l[1]}');
}

void main() {
  __vybeMain();
  __check('255:0');
}
