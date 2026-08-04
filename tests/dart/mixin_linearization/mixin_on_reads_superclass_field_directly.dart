// vybe-test: dart/mixin_linearization/mixin_on_reads_superclass_field_directly
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

class Vehicle {
  int wheels = 4;
}
mixin Inspect on Vehicle {
  int countWheels() {
    return wheels;
  }
}
class Truck extends Vehicle with Inspect {}
void __vybeMain() {
  __p(Truck().countWheels());
}

void main() {
  __vybeMain();
  __check('4');
}
