// vybe-test: dart/type_system/is_in_if
// origin: languages/dart/tests/dart/test_type_system.rs

void f(dynamic x) { if (x is String) { print(x.toUpperCase()); } }

void main() {}
