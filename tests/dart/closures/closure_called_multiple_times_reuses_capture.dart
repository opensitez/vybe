// vybe-test: dart/closures/closure_called_multiple_times_reuses_capture
// origin: languages/dart/tests/dart/test_closures.rs

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
  var base = 2;
  var addBase = (int n) => base + n;
  __p(addBase(1));
  __p(addBase(2));
  __p(addBase(3));
}

void main() {
  __vybeMain();
  __check('3\n4\n5');
}
