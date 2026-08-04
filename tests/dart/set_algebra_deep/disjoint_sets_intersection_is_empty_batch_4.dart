// vybe-test: dart/set_algebra_deep/disjoint_sets_intersection_is_empty_batch_4
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
  var a = {0,2,4,6,8};
  var b = {1,3,5,7,9};
  var c = a.intersection(b);
  __p(c.isEmpty);
  __p(c.length);
}

void main() {
  __vybeMain();
  __check('true\n0');
}
