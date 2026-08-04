// vybe-test: dart/bitwise_operators/bitwise_and_assign_updates_variable
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
  var x = 0xFF;
  x &= 0x0F;
  __p(x);
}

void main() {
  __vybeMain();
  __check('15');
}
