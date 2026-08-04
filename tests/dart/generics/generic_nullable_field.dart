// vybe-test: dart/generics/generic_nullable_field
// origin: languages/dart/tests/dart/test_generics.rs

class Cache<T> { T? _value; T? get value => _value; }

void main() {}
