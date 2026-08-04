// vybe-test: dart/extension_types/extension_type_int_add_two_meters
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
  Meters add(Meters other) {
    return Meters(value + other.value);
  }
}
void __vybeMain() {
  Meters a = Meters(2);
  Meters b = Meters(3);
  __p(a.add(b).value);
}

void main() {
  __vybeMain();
  __check('5');
}
