// vybe-test: dart/generics/generic_comparable
// origin: languages/dart/tests/dart/test_generics.rs

T max<T extends Comparable<T>>(T a, T b) => a.compareTo(b) > 0 ? a : b;

void main() {}
