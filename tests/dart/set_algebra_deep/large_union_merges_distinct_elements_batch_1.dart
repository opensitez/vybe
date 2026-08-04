// vybe-test: dart/set_algebra_deep/large_union_merges_distinct_elements_batch_1
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
  var a = {1,2,3,4,5};
  var b = {1,2,3,4,5};
  var c = a.union(b).toList()..sort();
  __p(c.length);
  __p(c.join(','));
}

void main() {
  __vybeMain();
  __check('5\n1,2,3,4,5');
}
