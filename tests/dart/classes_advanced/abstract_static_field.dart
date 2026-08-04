// vybe-test: dart/classes_advanced/abstract_static_field
// origin: languages/dart/tests/dart/test_classes_advanced.rs

abstract class Config {
  static const String version = '1.0';
  String get name;
}
class AppConfig extends Config {
  String get name => 'MyApp';
}

void main() {}
