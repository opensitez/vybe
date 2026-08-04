// vybe-test: dart/dart_typed_data_buffer_views/float32list_view
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
  bd.setFloat32(0, 3.5, Endian.host);
  final view = Float32List.view(bd.buffer, 0, 1);
  __p(view[0]);
}

void main() {
  __vybeMain();
  __check('3.5');
}
