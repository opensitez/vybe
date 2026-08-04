// vybe-test: dart/set_algebra_deep/large_intersection_keeps_shared_elements_batch_2
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
  var a = {1,3,5,7,9,11};
  var b = {2,3,5,7,11,13};
  var c = a.intersection(b).toList()..sort();
  __p(c.length);
  __p(c.join(','));
}

void main() {
  __vybeMain();
  __check('4\n3,5,7,11');
}
