// vybe-test: dart/set_algebra_deep/difference_after_union_preserves_left_only
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
  var a = {1, 2, 3};
  var b = {3, 4, 5};
  var u = a.union(b);
  var d = u.difference({4, 5});
  __p(d.length);
  __p(d.toList()..sort()..join(','));
}

void main() {
  __vybeMain();
  __check('3\n1,2,3');
}
