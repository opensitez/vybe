// vybe-test: dart/closures/closure_with_two_parameters_as_argument
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

int combine(int a, int b, int Function(int, int) fn) {
  return fn(a, b);
}
void __vybeMain() {
  __p(combine(3, 4, (x, y) => x * y));
}

void main() {
  __vybeMain();
  __check('12');
}
