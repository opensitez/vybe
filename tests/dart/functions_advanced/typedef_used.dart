// vybe-test: dart/functions_advanced/typedef_used
// origin: languages/dart/tests/dart/test_functions_advanced.rs

typedef Transformer = String Function(String); String apply(String s, Transformer t) { return t(s); }

void main() {}
