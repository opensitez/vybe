// vybe-test: dart/classes_advanced/multiple_mixins
// origin: languages/dart/tests/dart/test_classes_advanced.rs

mixin Flyable { void fly() { print('flying'); } }
mixin Swimmable { void swim() { print('swimming'); } }
class Duck with Flyable, Swimmable {
  String name;
  Duck(this.name);
}

void main() {}
