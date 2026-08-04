// vybe-test: dart/set_algebra_deep/large_union_merges_distinct_elements_batch_4
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
  var a = {0,2,4,6,8,10,12,14,16,18};
  var b = {1,3,5,7,9,11,13,15,17,19};
  var c = a.union(b).toList()..sort();
  __p(c.length);
  __p(c.join(','));
}

void main() {
  __vybeMain();
  __check('20\n0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19');
}
