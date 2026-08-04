// vybe-test: dart/functions_core/local_function_calls_sibling_local_function
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
  int double(int x) {
    return x * 2;
  }
  int quadruple(int x) {
    return double(double(x));
  }
  __p(quadruple(3));
}

void main() {
  __vybeMain();
  __check('12');
}
