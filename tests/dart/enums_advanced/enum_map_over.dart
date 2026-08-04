// vybe-test: dart/enums_advanced/enum_map_over
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Priority { low, medium, high }
void main() {
  var names = Priority.values.map((e) => e.name).toList();
}