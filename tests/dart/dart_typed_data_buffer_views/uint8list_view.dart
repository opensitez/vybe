// vybe-test: dart/dart_typed_data_buffer_views/uint8list_view
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
  final buffer = Uint8List.fromList([1, 2, 3, 4]).buffer;
  final view = Uint8List.view(buffer, 1, 2);
  // Damaged expectation repaired: the single interpolated print produces
  // "2:2" (measured, dart 3.10.4) — the original want "2\n2" assumed two
  // prints that were never written.
  __p('${view.length}:${view[0]}');
}

void main() {
  __vybeMain();
  __check('2:2');
}
