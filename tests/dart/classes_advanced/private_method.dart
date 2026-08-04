// vybe-test: dart/classes_advanced/private_method
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Parser { String _clean(String s) => s.trim(); String parse(String s) => _clean(s); }

void main() {}
