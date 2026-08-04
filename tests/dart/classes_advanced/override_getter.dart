// vybe-test: dart/classes_advanced/override_getter
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Base { int get value => 0; }
class Child extends Base {
  @override
  int get value => 42;
}

void main() {}
