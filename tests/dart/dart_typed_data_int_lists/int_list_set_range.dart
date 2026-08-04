// vybe-test: dart/dart_typed_data_int_lists/int_list_set_range
// origin: languages/dart/tests/dart/test_dart_typed_data_int_lists.rs

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
  final l1 = Uint8List(5);
  final l2 = Uint8List.fromList([1, 2, 3]);
  l1.setRange(1, 4, l2);
  __p('${l1[0]}:${l1[1]}:${l1[3]}:${l1[4]}');
}

void main() {
  __vybeMain();
  __check('0:1:3:0');
}
