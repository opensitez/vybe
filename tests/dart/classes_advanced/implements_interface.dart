// vybe-test: dart/classes_advanced/implements_interface
// origin: languages/dart/tests/dart/test_classes_advanced.rs

abstract class Printable { void print_(); } class Doc implements Printable { void print_() { print('doc'); } }

void main() {}
