// vybe-test: dart/null_safety_advanced/nullable_chain_method
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

class A { String? name; } void main() { var a = A(); var v = a.name?.toUpperCase(); }