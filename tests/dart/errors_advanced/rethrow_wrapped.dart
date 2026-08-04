// vybe-test: dart/errors_advanced/rethrow_wrapped
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void inner() { throw 'something'; }
void outer() {
  try {
    inner();
  } catch (e) {
    rethrow;
  }
}

void main() {}
