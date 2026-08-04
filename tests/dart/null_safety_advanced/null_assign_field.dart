// vybe-test: dart/null_safety_advanced/null_assign_field
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

class Cache {
  String? _value;
  String get value { _value ??= 'default'; return _value!; }
}

void main() {}
