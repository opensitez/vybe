// vybe-test: dart/generics/generic_extends_class
// origin: languages/dart/tests/dart/test_generics.rs

class Base<T> { T value; Base(this.value); } class Child<T> extends Base<T> { Child(T v) : super(v); }

void main() {}
