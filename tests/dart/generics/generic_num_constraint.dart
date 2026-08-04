// vybe-test: dart/generics/generic_num_constraint
// origin: languages/dart/tests/dart/test_generics.rs

T add<T extends num>(T a, T b) => (a + b) as T;

void main() {}
