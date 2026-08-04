// vybe-test: dart/errors_advanced/rethrow_basic
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void handle() {
  try {
    throw 'inner error';
  } catch (e) {
    rethrow;
  }
}

void main() {}
