// vybe-test: dart/type_system/nullable_param
// origin: languages/dart/tests/dart/test_type_system.rs

void greet(String? name) { print(name ?? 'guest'); }

void main() {}
