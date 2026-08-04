// vybe-test: dart/math_library/math_sin_cos_pythagorean_on_unit_circle
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
  var s = math.sin(math.pi / 4);
  var c = math.cos(math.pi / 4);
  __p((s * s + c * c) > 0.99);
}

void main() {
  __vybeMain();
  __check('true');
}
