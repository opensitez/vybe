// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_byte_data_view_reflects_original
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
  final bd = ByteData(4);
  final ubd = UnmodifiableByteDataView(bd);
  bd.setInt32(0, 12345);
  __p(ubd.getInt32(0));
}

void main() {
  __vybeMain();
  __check('12345');
}
