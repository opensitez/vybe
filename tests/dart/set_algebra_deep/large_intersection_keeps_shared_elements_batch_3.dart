// vybe-test: dart/set_algebra_deep/large_intersection_keeps_shared_elements_batch_3
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
  var a = {100,101,102,103,104,105,106,107,108,109};
  var b = {105,106,107,108,109,110,111,112,113,114};
  var c = a.intersection(b).toList()..sort();
  __p(c.length);
  __p(c.join(','));
}

void main() {
  __vybeMain();
  __check('5\n105,106,107,108,109');
}
