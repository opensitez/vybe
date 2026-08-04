// vybe-test: dart/functions_advanced/fn_as_arg
// origin: languages/dart/tests/dart/test_functions_advanced.rs

void apply(int x, int Function(int) fn) { print(fn(x)); }

void main() {}
