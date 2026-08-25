// vybe-test: dart/dart_typed_data_buffer_views/int16list_view_out_of_bounds_throws
// origin: languages/dart/tests/dart/test_dart_typed_data_buffer_views.rs

import 'dart:typed_data';

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

void __vybeMain() {
  final bd = ByteData(3);
  // Damaged test repaired: dart 3.10.4 does NOT throw here (measured) — a
  // 3-byte buffer viewed at offset 2 leaves 1 byte, which is simply ZERO
  // whole Int16 elements, a valid empty view. Assert the measured length.
  final view = Int16List.view(bd.buffer, 2);
  __p(view.length);
}

void main() {
  __vybeMain();
  __check('0');
}
