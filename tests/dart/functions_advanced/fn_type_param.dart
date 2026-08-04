// vybe-test: dart/functions_advanced/fn_type_param
// origin: languages/dart/tests/dart/test_functions_advanced.rs

T apply<T>(T val, T Function(T) fn) => fn(val);

void main() {}
