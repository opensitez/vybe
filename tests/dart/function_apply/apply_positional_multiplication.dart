// vybe-test: dart/function_apply/apply_positional_multiplication
// origin: languages/dart/tests/dart/test_function_apply.rs

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

int multiply(int a, int b) {
  return a * b;
}
void __vybeMain() {
  __p(Function.apply(multiply, [6, 7]));
}

void main() {
  __vybeMain();
  __check('42');
}
