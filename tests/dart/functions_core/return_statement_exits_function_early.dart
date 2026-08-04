// vybe-test: dart/functions_core/return_statement_exits_function_early
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

int absDiff(int a, int b) {
  if (a >= b) {
    return a - b;
  }
  return b - a;
}
void __vybeMain() {
  __p(absDiff(5, 9));
  __p(absDiff(9, 5));
}

void main() {
  __vybeMain();
  __check('4\n4');
}
