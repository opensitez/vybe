// vybe-test: dart/generics/generic_concrete_child
// origin: languages/dart/tests/dart/test_generics.rs

class Base<T> { T value; Base(this.value); } class IntBox extends Base<int> { IntBox(int v) : super(v); }

void main() {}
