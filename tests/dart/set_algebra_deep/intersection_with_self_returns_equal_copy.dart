// vybe-test: dart/set_algebra_deep/intersection_with_self_returns_equal_copy
// origin: languages/dart/tests/dart/test_set_algebra_deep.rs

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
  var s = {10, 20, 30, 40, 50};
  var i = s.intersection(s);
  __p(i.length);
  __p(i == s);
}

void main() {
  __vybeMain();
  __check('5\ntrue');
}
