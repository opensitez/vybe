// vybe-test: dart/features/inherited_method
// origin: languages/dart/tests/dart/test_features.rs

class A { int value() { return 42; } } class B extends A {} void main() { var b = B(); print(b.value()); }