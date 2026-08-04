// vybe-test: dart/math_library/math_pow_composes_with_sqrt
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
  __p(math.sqrt(math.pow(6, 2)));
}

void main() {
  __vybeMain();
  __check('6.0');
}
