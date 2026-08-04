// vybe-test: dart/dart_typed_data_buffer_views/buffer_view_modifies_original
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
  final list = Uint8List(4);
  final view = Uint32List.view(list.buffer);
  view[0] = 0xFFFFFFFF;
  __p('${list[0]}:${list[1]}:${list[2]}:${list[3]}');
}

void main() {
  __vybeMain();
  __check('255:255:255:255');
}
