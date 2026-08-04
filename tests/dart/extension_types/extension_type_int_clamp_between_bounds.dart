// vybe-test: dart/extension_types/extension_type_int_clamp_between_bounds
// origin: languages/dart/tests/dart/test_extension_types.rs

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

extension type Percent(int value) {
  Percent clamped() {
    if (value < 0) return Percent(0);
    if (value > 100) return Percent(100);
    return Percent(value);
  }
}
void __vybeMain() {
  Percent p = Percent(150);
  __p(p.clamped().value);
}

void main() {
  __vybeMain();
  __check('100');
}
