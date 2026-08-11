// vybe-test: dart/set_algebra_deep/union_with_disjoint_sets_keeps_all_members
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
  var a = {0, 2, 4};
  var b = {1, 3, 5};
  var r = a.union(b).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('6\n0,1,2,3,4,5');
}
