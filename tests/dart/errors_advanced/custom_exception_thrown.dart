// vybe-test: dart/errors_advanced/custom_exception_thrown
// origin: languages/dart/tests/dart/test_errors_advanced.rs

class AppException implements Exception {
  final String message;
  AppException(this.message);
}
void risky() { throw AppException('Something went wrong'); }

void main() {}
