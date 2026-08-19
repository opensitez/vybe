// vybe-test: dart/dart_typed_data_buffer_views/float32x4list_view
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
  final bd = ByteData(16);
  bd.setFloat32(0, 1.0, Endian.host);
  bd.setFloat32(4, 2.0, Endian.host);
  bd.setFloat32(8, 3.0, Endian.host);
  bd.setFloat32(12, 4.0, Endian.host);
  final view = Float32x4List.view(bd.buffer);
  __p(view[0].z);
}

void main() {
  __vybeMain();
  __check('3.0');
}
