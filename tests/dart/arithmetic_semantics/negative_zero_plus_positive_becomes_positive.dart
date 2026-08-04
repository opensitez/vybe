// vybe-test: dart/arithmetic_semantics/negative_zero_plus_positive_becomes_positive
// origin: languages/dart/tests/dart/test_arithmetic_semantics.rs

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
  __p(-0.0 + 1.0);
}

void main() {
  __vybeMain();
  __check('1.0');
}
