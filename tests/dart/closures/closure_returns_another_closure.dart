// vybe-test: dart/closures/closure_returns_another_closure
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

Function makeChain(int a) {
  return (int b) {
    return (int c) => a + b + c;
  };
}
void __vybeMain() {
  var step = makeChain(1)(2);
  __p(step(3));
}

void main() {
  __vybeMain();
  __check('6');
}
