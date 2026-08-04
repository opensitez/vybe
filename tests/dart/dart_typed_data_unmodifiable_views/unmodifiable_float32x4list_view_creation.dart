// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_float32x4list_view_creation
// origin: languages/dart/tests/dart/test_dart_typed_data_unmodifiable_views.rs

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
  final l = Float32x4List.fromList([Float32x4(1.0, 2.0, 3.0, 4.0)]);
  final ul = UnmodifiableFloat32x4ListView(l);
  __p(ul[0].z);
}

void main() {
  __vybeMain();
  __check('3.0');
}
