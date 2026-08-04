// vybe-test: dart/unmodifiable_collections/unmod_set_contains_checks_membership
// origin: languages/dart/tests/dart/test_unmodifiable_collections.rs

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
  var frozen = Set.unmodifiable({'x', 'y'});
  __p(frozen.contains('x'));
  __p(frozen.contains('z'));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
