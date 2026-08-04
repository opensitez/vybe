// vybe-test: dart/bitwise_operators/right_shift_divides_by_power_of_two
// origin: languages/dart/tests/dart/test_bitwise_operators.rs

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
  __p(64 >> 2);
}

void main() {
  __vybeMain();
  __check('16');
}
