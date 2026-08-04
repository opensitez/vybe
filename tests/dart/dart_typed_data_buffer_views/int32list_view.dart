// vybe-test: dart/dart_typed_data_buffer_views/int32list_view
// origin: languages/dart/tests/dart/test_dart_typed_data_buffer_views.rs

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
  final bd = ByteData(8);
  bd.setInt32(4, -123456, Endian.host);
  final view = Int32List.view(bd.buffer, 4, 1);
  __p(view[0]);
}

void main() {
  __vybeMain();
  __check('-123456');
}
