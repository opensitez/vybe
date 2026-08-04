// vybe-test: dart/generics/generic_extends
// origin: languages/dart/tests/dart/test_generics.rs

class Container<T extends Object> { T value; Container(this.value); }

void main() {}
