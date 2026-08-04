// vybe-test: dart/set_core/set_of_copies_from_existing_set
// origin: languages/dart/tests/dart/test_set_core.rs

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
  var src = {5, 6, 7};
  var copy = Set<int>.of(src);
  __p(copy.length);
  __p(copy.contains(6));
}

void main() {
  __vybeMain();
  __check('3\ntrue');
}
