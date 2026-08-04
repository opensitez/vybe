// vybe-test: dart/functions_core/function_returns_another_function
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

int Function(int) makeAdder(int n) {
  return (x) => x + n;
}
void __vybeMain() {
  var add10 = makeAdder(10);
  __p(add10(5));
}

void main() {
  __vybeMain();
  __check('15');
}
