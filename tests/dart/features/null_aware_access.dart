// vybe-test: dart/features/null_aware_access
// origin: languages/dart/tests/dart/test_features.rs

class A { int x = 1; } void main() { var a = null; var b = a?.x; }