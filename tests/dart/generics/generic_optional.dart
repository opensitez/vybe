// vybe-test: dart/generics/generic_optional
// origin: languages/dart/tests/dart/test_generics.rs

class Maybe<T> { T? value; Maybe([this.value]); bool get hasValue => value != null; }

void main() {}
