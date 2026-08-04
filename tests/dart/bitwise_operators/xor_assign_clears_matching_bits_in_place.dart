// vybe-test: dart/bitwise_operators/xor_assign_clears_matching_bits_in_place
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
  var x = 0b1010;
  x ^= 0b1100;
  __p(x);
}

void main() {
  __vybeMain();
  __check('6');
}
