// vybe-test: dart/const_final/final_late_const_in_class
// origin: languages/dart/tests/dart/test_const_final.rs

class Settings {
  static const String version = '1.0';
  final String name;
  late String description;
  Settings(this.name);
}

void main() {}
