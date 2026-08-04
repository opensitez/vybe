// vybe-test: dart/classes_advanced/static_counter
// origin: languages/dart/tests/dart/test_classes_advanced.rs

class Counter {
  static int _count = 0;
  static void increment() { _count++; }
  static int get count => _count;
}

void main() {}
