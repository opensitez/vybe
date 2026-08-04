// vybe-test: dart/null_safety_advanced/null_guard_return
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

void process(String? s) {
  if (s == null) return;
  print(s.toUpperCase());
}

void main() {}
