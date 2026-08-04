// vybe-test: dart/type_system/nullable_chain
// origin: languages/dart/tests/dart/test_type_system.rs

class A { int x = 1; } void main() { A? a = null; var v = a?.x; }