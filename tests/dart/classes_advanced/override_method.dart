// vybe-test: dart/classes_advanced/override_method
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Animal { String speak() => 'animal sound'; }
class Cat extends Animal {
  @override
  String speak() => 'meow';
}

void main() {}
