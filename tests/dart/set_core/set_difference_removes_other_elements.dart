// vybe-test: dart/set_core/set_difference_removes_other_elements
// origin: languages/dart/tests/dart/test_set_core.rs

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
  var b = {2};
  __p(a.difference(b).toList().join(','));
  __p(a.difference(b).length);
}

void main() {
  __vybeMain();
  __check('1,3\n2');
}
