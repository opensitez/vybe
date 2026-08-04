// vybe-test: dart/dart_typed_data_byte_data_endianness/byte_data_view_offset
// origin: languages/dart/tests/dart/test_dart_typed_data_byte_data_endianness.rs

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
  final list = Uint8List.fromList([10, 20, 30, 40]);
  final bd = ByteData.view(list.buffer, 2);
  __p(bd.lengthInBytes);
  __p(bd.getUint8(0));
}

void main() {
  __vybeMain();
  __check('2\n30');
}
