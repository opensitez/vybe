// vybe-test: dart/dart_typed_data_buffer_views/int64list_view
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
  final bd = ByteData(8);
  bd.setInt64(0, -123456789, Endian.host);
  final view = Int64List.view(bd.buffer);
  __p(view[0]);
}

void main() {
  __vybeMain();
  __check('-123456789');
}
