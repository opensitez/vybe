// vybe-test: dart/functions_core/arrow_function_assigned_to_variable_and_called
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
  bool Function(int) isPositive = (n) => n > 0;
  __p(isPositive(1));
  __p(isPositive(-1));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
