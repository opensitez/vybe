// vybe-test: dart/errors_advanced/on_with_catch_fallback
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() {
  try {
    throw Exception('any');
  } on FormatException {
    print('format');
  } catch (e) {
    print('fallback');
  }
}