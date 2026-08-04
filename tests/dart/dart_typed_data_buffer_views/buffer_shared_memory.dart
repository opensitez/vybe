// vybe-test: dart/dart_typed_data_buffer_views/buffer_shared_memory
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
  final list1 = Uint8List(4);
  final list2 = Uint16List.view(list1.buffer);
  list1[0] = 0xFF;
  list1[1] = 0x00;
  // If host is little endian, 0x00FF = 255
  // If host is big endian, 0xFF00 = 65280
  final v = list2[0];
  __p(v == 255 || v == 65280);
}

void main() {
  __vybeMain();
  __check('true');
}
