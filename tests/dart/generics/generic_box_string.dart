// vybe-test: dart/generics/generic_box_string
// origin: languages/dart/tests/dart/test_generics.rs

class Box<T> { T value; Box(this.value); } void main() { var b = Box('hello'); }