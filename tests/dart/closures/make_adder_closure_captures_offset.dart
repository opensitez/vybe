// vybe-test: dart/closures/make_adder_closure_captures_offset
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

int Function(int) makeAdder(int n) {
  return (x) => x + n;
}
void __vybeMain() {
  var add5 = makeAdder(5);
  __p(add5(10));
}

void main() {
  __vybeMain();
  __check('15');
}
