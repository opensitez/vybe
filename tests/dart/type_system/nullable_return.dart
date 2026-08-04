// vybe-test: dart/type_system/nullable_return
// origin: languages/dart/tests/dart/test_type_system.rs

String? findName(bool found) { return found ? 'Alice' : null; }

void main() {}
