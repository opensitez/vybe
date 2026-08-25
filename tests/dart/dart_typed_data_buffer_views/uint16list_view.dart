// vybe-test: dart/dart_typed_data_buffer_views/uint16list_view
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
  final bd = ByteData(4);
  bd.setUint16(0, 500, Endian.host);
  bd.setUint16(2, 1000, Endian.host);
  final view = Uint16List.view(bd.buffer);
  // Damaged expectation repaired: the single interpolated print produces
  // "2:500:1000" (measured, dart 3.10.4) — the original want assumed three
  // prints that were never written.
  __p('${view.length}:${view[0]}:${view[1]}');
}

void main() {
  __vybeMain();
  __check('2:500:1000');
}
