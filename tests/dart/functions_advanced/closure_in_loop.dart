// vybe-test: dart/functions_advanced/closure_in_loop
// origin: languages/dart/tests/dart/test_functions_advanced.rs

void main() { var fns = []; for (var i = 0; i < 3; i++) { fns.add(() => i); } }