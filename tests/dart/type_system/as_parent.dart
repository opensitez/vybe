// vybe-test: dart/type_system/as_parent
// origin: languages/dart/tests/dart/test_type_system.rs

class Animal { String name = 'A'; } class Dog extends Animal {} void main() { Dog d = Dog(); Animal a = d as Animal; }