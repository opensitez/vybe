// vybe-test: dart/dart_typed_data_buffer_views/view_offset_in_bytes
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
  final list = Uint8List(10);
  final view = Uint32List.view(list.buffer, 4, 1);
  __p(view.offsetInBytes);
}

void main() {
  __vybeMain();
  __check('4');
}
