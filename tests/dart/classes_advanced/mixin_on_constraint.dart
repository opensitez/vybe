// vybe-test: dart/classes_advanced/mixin_on_constraint
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Animal { String name; Animal(this.name); }
mixin Domestic on Animal { String get owner => 'human'; }
class Dog extends Animal with Domestic { Dog(String n) : super(n); }

void main() {}
