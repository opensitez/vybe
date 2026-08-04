// vybe-test: dart/features/mixin_method
// origin: languages/dart/tests/dart/test_features.rs

class Printable { String describe() { return 'printable'; } } class Foo with Printable {}

void main() {}
