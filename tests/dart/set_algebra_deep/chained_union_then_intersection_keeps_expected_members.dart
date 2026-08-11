// vybe-test: dart/set_algebra_deep/chained_union_then_intersection_keeps_expected_members
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
  var a = {1, 2, 3, 4};
  var b = {3, 4, 5, 6};
  var c = {2, 3, 4, 7};
  var r = a.union(b).intersection(c).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('3\n2,3,4');
}
