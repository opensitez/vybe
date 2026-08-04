// vybe-test: dart/closures/closure_compose_two_unary_functions
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

int compose(int x, int Function(int) f, int Function(int) g) {
  return f(g(x));
}
void __vybeMain() {
  __p(compose(3, (n) => n + 1, (n) => n * 2));
}

void main() {
  __vybeMain();
  __check('7');
}
