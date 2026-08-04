// vybe-test: dart/dart_apis/class_inheritance
// origin: languages/dart/tests/dart/test_dart_apis.rs

class Animal { String name; Animal(this.name); } class Dog extends Animal { Dog(String name) : super(name); String speak() => name + ' barks'; }

void main() {}
