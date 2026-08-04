// vybe-test: dart/double_special_values/positive_division_by_zero_yields_infinity
// origin: languages/dart/tests/dart/test_double_special_values.rs

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
  __p(1.0 / 0.0);
}

void main() {
  __vybeMain();
  __check('Infinity');
}
