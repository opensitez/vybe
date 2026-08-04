// vybe-test: dart/set_algebra_deep/mutual_contains_all_implies_equal_large_sets
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
  var a = {for (var i = 0; i < 12; i++) i * 2};
  var b = {0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22};
  __p(a.containsAll(b));
  __p(b.containsAll(a));
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true\ntrue\ntrue');
}
