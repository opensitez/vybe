// vybe-test: dart/type_system/non_null_in_method
// origin: languages/dart/tests/dart/test_type_system.rs

class A { int? x = 5; int get val => x!; }

void main() {}
