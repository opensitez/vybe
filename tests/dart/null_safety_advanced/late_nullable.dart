// vybe-test: dart/null_safety_advanced/late_nullable
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

late String? name; void main() { name = null; print(name ?? 'empty'); }