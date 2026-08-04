// vybe-test: dart/logical_operators/short_circuit_or_skips_rhs_increment
// origin: languages/dart/tests/dart/test_logical_operators.rs

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
  var steps = 0;
  true || (steps = steps + 1) == 1;
  __p(steps);
}

void main() {
  __vybeMain();
  __check('0');
}
