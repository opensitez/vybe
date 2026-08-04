// vybe-test: dart/generics/generic_fn_two
// origin: languages/dart/tests/dart/test_generics.rs

B transform<A, B>(A val, B Function(A) fn) => fn(val);

void main() {}
