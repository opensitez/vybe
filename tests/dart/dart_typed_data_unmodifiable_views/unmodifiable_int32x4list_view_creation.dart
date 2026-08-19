// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_int32x4list_view_creation
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
  final l = Int32x4List.fromList([Int32x4(1, 2, 3, 4)]);
  final ul = UnmodifiableInt32x4ListView(l);
  __p(ul[0].w);
}

void main() {
  __vybeMain();
  __check('4');
}
