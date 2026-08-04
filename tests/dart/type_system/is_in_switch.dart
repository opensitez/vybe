// vybe-test: dart/type_system/is_in_switch
// origin: languages/dart/tests/dart/test_type_system.rs

void describe(dynamic x) {
  if (x is int) {
    print('int: $x');
  } else if (x is String) {
    print('string: $x');
  } else {
    print('other');
  }
}

void main() {}
