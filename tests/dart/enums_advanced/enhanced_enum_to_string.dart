// vybe-test: dart/enums_advanced/enhanced_enum_to_string
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Color {
  red, green, blue;
  String describe() => 'Color.$name';
}

void main() {}
