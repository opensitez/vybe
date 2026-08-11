// vybe-test: dart/set_algebra_deep/union_merges_overlap_and_preserves_unique_elements
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
  var a = {1, 2, 3, 4, 5};
  var b = {4, 5, 6, 7};
  var r = a.union(b).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('7\n1,2,3,4,5,6,7');
}
