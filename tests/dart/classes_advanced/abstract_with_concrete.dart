// vybe-test: dart/classes_advanced/abstract_with_concrete
// origin: languages/dart/tests/dart/test_classes_advanced.rs

abstract class Vehicle {
  String name;
  Vehicle(this.name);
  void fuel() { print('fueling $name'); }
  void drive();
}
class Car extends Vehicle {
  Car(String n) : super(n);
  void drive() { print('$name driving'); }
}

void main() {}
