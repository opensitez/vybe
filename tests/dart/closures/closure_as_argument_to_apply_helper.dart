// vybe-test: dart/closures/closure_as_argument_to_apply_helper
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

int apply(int x, int Function(int) fn) {
  return fn(x);
}
void __vybeMain() {
  __p(apply(5, (n) => n + 1));
}

void main() {
  __vybeMain();
  __check('6');
}
