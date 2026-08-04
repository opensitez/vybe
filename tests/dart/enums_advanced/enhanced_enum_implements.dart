// vybe-test: dart/enums_advanced/enhanced_enum_implements
// origin: languages/dart/tests/dart/test_enums_advanced.rs

abstract class Describable { String describe(); }
enum Status implements Describable {
  active, inactive;
  String describe() => 'Status is $name';
}

void main() {}
