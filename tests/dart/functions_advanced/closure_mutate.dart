// vybe-test: dart/functions_advanced/closure_mutate
// origin: languages/dart/tests/dart/test_functions_advanced.rs

void main() { var count = 0; var inc = () { count++; }; inc(); inc(); print(count); }