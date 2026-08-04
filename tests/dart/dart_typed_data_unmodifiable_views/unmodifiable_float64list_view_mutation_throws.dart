// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_float64list_view_mutation_throws
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
  final l = Float64List.fromList([1.1]);
  final ul = UnmodifiableFloat64ListView(l);
  try {
    ul[0] = 2.2;
  } on UnsupportedError {
    __p('UnsupportedError thrown');
  }
}

void main() {
  __vybeMain();
  __check('UnsupportedError thrown');
}
