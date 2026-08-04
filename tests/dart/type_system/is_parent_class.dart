// vybe-test: dart/type_system/is_parent_class
// origin: languages/dart/tests/dart/test_type_system.rs

class Animal {} class Dog extends Animal {} void main() { var d = Dog(); var b = d is Animal; }