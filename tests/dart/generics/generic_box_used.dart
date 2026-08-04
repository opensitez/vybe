// vybe-test: dart/generics/generic_box_used
// origin: languages/dart/tests/dart/test_generics.rs

class Box<T> { T value; Box(this.value); } void main() { var b = Box(42); print(b.value); }