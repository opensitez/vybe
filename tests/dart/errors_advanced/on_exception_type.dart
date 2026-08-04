// vybe-test: dart/errors_advanced/on_exception_type
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() {
  try {
    throw FormatException('bad format');
  } on FormatException catch (e) {
    print('format error');
  }
}