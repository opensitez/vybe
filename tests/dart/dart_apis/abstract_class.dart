// vybe-test: dart/dart_apis/abstract_class
// origin: languages/dart/tests/dart/test_dart_apis.rs

abstract class Shape { double area(); } class Circle extends Shape { double r; Circle(this.r); double area() => 3.14 * r * r; }

void main() {}
