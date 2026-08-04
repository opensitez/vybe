// vybe-test: dart/errors_advanced/custom_error
// origin: languages/dart/tests/dart/test_errors_advanced.rs

class StateError extends Error { final String msg; StateError(this.msg); }

void main() {}
