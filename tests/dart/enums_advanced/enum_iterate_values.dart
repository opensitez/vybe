// vybe-test: dart/enums_advanced/enum_iterate_values
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Color { red, green, blue }
void main() {
  for (var c in Color.values) {
    print(c.name);
  }
}