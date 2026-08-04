// vybe-test: dart/logical_operators/short_circuit_and_prints_only_lhs_side_effect
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
  bool mark(String label) {
    __p(label);
    return true;
  }
  false && mark('rhs');
  __p('end');
}

void main() {
  __vybeMain();
  __check('end');
}
