// vybe-test: dart/extension_types/extension_type_int_multiply_by_scalar
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

extension type Meters(int value) {
  Meters scale(int factor) {
    return Meters(value * factor);
  }
}
void __vybeMain() {
  Meters m = Meters(4);
  __p(m.scale(3).value);
}

void main() {
  __vybeMain();
  __check('12');
}
