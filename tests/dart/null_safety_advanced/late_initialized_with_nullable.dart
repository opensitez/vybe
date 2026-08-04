// vybe-test: dart/null_safety_advanced/late_initialized_with_nullable
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

class Loader {
  late String? data;
  void load(bool found) { data = found ? 'result' : null; }
}

void main() {}
