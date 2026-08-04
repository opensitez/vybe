// vybe-test: dart/closures/closure_wrapper_delegates_to_captured_fn
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
  int doubleIt(int x) => x * 2;
  var wrap = (int n) => doubleIt(n) + 1;
  __p(wrap(4));
}

void main() {
  __vybeMain();
  __check('9');
}
