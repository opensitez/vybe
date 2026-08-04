// vybe-test: dart/generics/generic_nullable_param
// origin: languages/dart/tests/dart/test_generics.rs

T? tryParse<T>(String s, T? Function(String) parser) => parser(s);

void main() {}
