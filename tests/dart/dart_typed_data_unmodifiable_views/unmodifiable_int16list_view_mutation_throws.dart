// vybe-test: dart/dart_typed_data_unmodifiable_views/unmodifiable_int16list_view_mutation_throws
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
  final l = Int16List.fromList([-1000]);
  final ul = UnmodifiableInt16ListView(l);
  try {
    ul[0] = 0;
  } on UnsupportedError {
    __p('UnsupportedError thrown');
  }
}

void main() {
  __vybeMain();
  __check('UnsupportedError thrown');
}
