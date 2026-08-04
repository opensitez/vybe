// vybe-test: dart/generics/generic_two_type_params
// origin: languages/dart/tests/dart/test_generics.rs

class Result<T, E> { T? success; E? error; Result.ok(this.success); Result.err(this.error); }

void main() {}
