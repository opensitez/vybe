// vybe-test: dart/features/super_method_call
// origin: languages/dart/tests/dart/test_features.rs

class A { String greet() { return 'hello'; } } class B extends A { String greet() { return super.greet() + ' world'; } }

void main() {}
