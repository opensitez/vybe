// vybe-test: dart/functions_core/local_recursive_function_computes_factorial
// origin: languages/dart/tests/dart/test_functions_core.rs

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
  int fact(int n) {
    if (n <= 1) {
      return 1;
    }
    return n * fact(n - 1);
  }
  __p(fact(6));
}

void main() {
  __vybeMain();
  __check('720');
}
