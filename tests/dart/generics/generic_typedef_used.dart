// vybe-test: dart/generics/generic_typedef_used
// origin: languages/dart/tests/dart/test_generics.rs

typedef Predicate<T> = bool Function(T); bool check<T>(T val, Predicate<T> pred) => pred(val);

void main() {}
