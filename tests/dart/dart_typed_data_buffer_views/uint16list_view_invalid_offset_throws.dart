// vybe-test: dart/dart_typed_data_buffer_views/uint16list_view_invalid_offset_throws
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
  try {
    // Offset must be a multiple of 2 for Uint16List
    Uint16List.view(bd.buffer, 1);
  } on ArgumentError {
    __p('ArgumentError thrown');
  }
}

void main() {
  __vybeMain();
  __check('ArgumentError thrown');
}
