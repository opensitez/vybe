// vybe-test: dart/functions_advanced/recursive_factorial_result
// origin: languages/dart/tests/dart/test_functions_advanced.rs

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

int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }
void __vybeMain() { __p(fact(5)); }

void main() {
  __vybeMain();
  __check('120');
}
