// vybe-test: dart/abstract_members/abstract_class_constructor_in_subclass
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Vehicle {
  String make;
  Vehicle(this.make);
  void drive();
}
class Car extends Vehicle {
  Car(String m) : super(m);
  void drive() {
    __p(make);
  }
}
void __vybeMain() {
  Car('Vybe').drive();
}

void main() {
  __vybeMain();
  __check('Vybe');
}
