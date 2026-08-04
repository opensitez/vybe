// vybe-test: dart/null_safety_advanced/assert_non_null_method
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

class A { int? val = 5; void f() { print(val!); } }

void main() {}
