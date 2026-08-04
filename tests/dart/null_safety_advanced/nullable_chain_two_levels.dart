// vybe-test: dart/null_safety_advanced/nullable_chain_two_levels
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

class B { int x = 1; } class A { B? b; } void main() { var a = A(); var v = a.b?.x; }