// vybe-test: dart/functions_advanced/recursive_factorial
// origin: languages/dart/tests/dart/test_functions_advanced.rs

int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }

void main() {}
