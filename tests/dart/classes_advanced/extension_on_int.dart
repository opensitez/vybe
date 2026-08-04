// vybe-test: dart/classes_advanced/extension_on_int
// origin: languages/dart/tests/dart/test_classes_advanced.rs

extension IntExt on int { bool get isEven => this % 2 == 0; }

void main() {}
