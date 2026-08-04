// vybe-test: dart/functions_advanced/closure_capture
// origin: languages/dart/tests/dart/test_functions_advanced.rs

void main() { var x = 10; var fn = () => x * 2; print(fn()); }