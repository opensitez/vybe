// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_byte_data_view_creation
// origin: languages/dart/tests/dart/test_dart_typed_data_unmodifiable_views.rs

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
  final l = Uint8List.fromList([1, 2, 3, 4]);
  final bd = ByteData.view(l.buffer);
  final ubd = UnmodifiableByteDataView(bd);
  __p(ubd.getUint8(1));
}

void main() {
  __vybeMain();
  __check('2');
}
