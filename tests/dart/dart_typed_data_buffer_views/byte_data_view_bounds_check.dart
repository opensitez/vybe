// vybe-test: dart/dart_typed_data_buffer_views/byte_data_view_bounds_check
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
  final buffer = Uint8List(5).buffer;
  try {
    ByteData.view(buffer, 6);
  } on RangeError {
    __p('RangeError thrown');
  } catch(e) {
    __p('ArgumentError thrown'); // Depending on impl
  }
}

void main() {
  __vybeMain();
  __check('RangeError thrown');
}
