// vybe-test: dart/features/class_inheritance
// origin: languages/dart/tests/dart/test_features.rs

class Animal { String name; Animal(this.name); } class Dog extends Animal { Dog(String n) : super(n); }

void main() {}
