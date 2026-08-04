// vybe-test: dart/classes_advanced/static_factory_pattern
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Singleton {
  static Singleton? _instance;
  static Singleton get instance {
    _instance ??= Singleton._();
    return _instance!;
  }
  Singleton._();
}

void main() {}
