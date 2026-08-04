// vybe-test: dart/errors_advanced/custom_exception
// origin: languages/dart/tests/dart/test_errors_advanced.rs

class AppException implements Exception { final String message; AppException(this.message); }

void main() {}
