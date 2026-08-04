// vybe-test: dart/features/any_every
// origin: languages/dart/tests/dart/test_features.rs

void main() { var h = [1,2,3].any((e) => e > 2); var a = [1,2,3].every((e) => e > 0); }