// vybe-test: dart/bitwise_operators/left_shift_assign_multiplies_in_place
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
  var x = 1;
  x <<= 3;
  __p(x);
}

void main() {
  __vybeMain();
  __check('8');
}
