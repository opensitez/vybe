// vybe-test: dart/math_library/double_abs_on_negative_integer_as_double
// origin: languages/dart/tests/dart/test_math_library.rs

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
  __p((-42.0).abs());
}

void main() {
  __vybeMain();
  __check('42.0');
}
