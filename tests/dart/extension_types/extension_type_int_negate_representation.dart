// vybe-test: dart/extension_types/extension_type_int_negate_representation
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

extension type Temperature(int celsius) {
  Temperature negate() {
    return Temperature(-celsius);
  }
}
void __vybeMain() {
  Temperature t = Temperature(20);
  __p(t.negate().celsius);
}

void main() {
  __vybeMain();
  __check('-20');
}
