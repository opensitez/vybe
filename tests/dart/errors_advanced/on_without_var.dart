// vybe-test: dart/errors_advanced/on_without_var
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() {
  try {
    throw FormatException('bad');
  } on FormatException {
    print('caught format');
  }
}