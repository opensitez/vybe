// vybe-test: dart/functions_core/top_level_function_returns_boolean_comparison
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

bool isEven(int n) {
  return n % 2 == 0;
}
void __vybeMain() {
  __p(isEven(4));
  __p(isEven(5));
}

void main() {
  __vybeMain();
  __check('true\nfalse');
}
