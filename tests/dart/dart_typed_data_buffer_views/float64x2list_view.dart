// vybe-test: dart/dart_typed_data_buffer_views/float64x2list_view
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
  final bd = ByteData(16);
  bd.setFloat64(0, 10.0, Endian.host);
  bd.setFloat64(8, 20.0, Endian.host);
  final view = Float64x2List.view(bd.buffer);
  __p('${view[0].x}:${view[0].y}');
}

void main() {
  __vybeMain();
  __check('10.0:20.0');
}
