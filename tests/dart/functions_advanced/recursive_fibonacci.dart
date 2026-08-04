// vybe-test: dart/functions_advanced/recursive_fibonacci
// origin: languages/dart/tests/dart/test_functions_advanced.rs

int fib(int n) { return n <= 1 ? n : fib(n - 1) + fib(n - 2); }

void main() {}
