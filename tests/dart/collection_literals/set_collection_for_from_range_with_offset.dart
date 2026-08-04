// vybe-test: dart/collection_literals/set_collection_for_from_range_with_offset
// origin: languages/dart/tests/dart/test_collection_literals.rs

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
  var s = {for (var i = 0; i < 3; i++) i + 10};
  __p(s.contains(12));
  __p(s.length);
}

void main() {
  __vybeMain();
  __check('true\n3');
}
