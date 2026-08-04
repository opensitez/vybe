// vybe-test: dart/class_modifiers/base_class_instantiation_and_method
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

base class Vehicle {
  String kind() {
    return 'vehicle';
  }
}
void __vybeMain() {
  __p(Vehicle().kind());
}

void main() {
  __vybeMain();
  __check('vehicle');
}
