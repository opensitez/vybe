// vybe-test: dart/enums_advanced/enhanced_enum_with_method
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Day {
  monday, tuesday, wednesday, thursday, friday, saturday, sunday;

  bool get isWeekend => this == Day.saturday || this == Day.sunday;
}

void main() {}
