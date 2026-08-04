// vybe-test: dart/covariant_keyword/covariant_param_two_level_hierarchy
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Vehicle {}
class Car extends Vehicle {}
class Garage {
  void park(Vehicle v) {}
}
class CarGarage extends Garage {
  @override
  void park(covariant Car c) {}
}
void __vybeMain() {
  CarGarage().park(Car());
  __p(1);
}

void main() {
  __vybeMain();
  __check('1');
}
