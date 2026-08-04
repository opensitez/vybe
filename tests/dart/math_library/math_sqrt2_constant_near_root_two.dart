// vybe-test: dart/math_library/math_sqrt2_constant_near_root_two
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
  __p(math.sqrt2 > 1.414 && math.sqrt2 < 1.415);
}

void main() {
  __vybeMain();
  __check('true');
}
