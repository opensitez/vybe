// vybe-test: dart/functions_advanced/fn_return_fn_result
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

Function makeAdder(int n) { return (int x) => x + n; }
void __vybeMain() { var add5 = makeAdder(5); __p(add5(3)); }

void main() {
  __vybeMain();
  __check('8');
}
