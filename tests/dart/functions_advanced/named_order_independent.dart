// vybe-test: dart/functions_advanced/named_order_independent
// origin: languages/dart/tests/dart/test_functions_advanced.rs

void f({int a = 0, int b = 0}) {} void main() { f(b: 2, a: 1); }