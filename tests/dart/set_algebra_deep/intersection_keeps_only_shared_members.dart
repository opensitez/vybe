// vybe-test: dart/set_algebra_deep/intersection_keeps_only_shared_members
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
  var a = {1, 2, 3, 4, 5, 6};
  var b = {4, 5, 6, 7, 8};
  var r = a.intersection(b).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('3\n4,5,6');
}
