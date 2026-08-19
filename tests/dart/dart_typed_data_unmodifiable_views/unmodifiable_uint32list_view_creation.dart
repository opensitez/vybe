// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_uint32list_view_creation
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
  final l = Uint32List.fromList([4000000000]);
  final ul = UnmodifiableUint32ListView(l);
  __p(ul[0]);
}

void main() {
  __vybeMain();
  __check('4000000000');
}
