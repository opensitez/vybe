// vybe-test: dart/errors_advanced/finally_no_catch
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void risky() {
  try {
    var x = 1;
  } finally {
    print('done');
  }
}

void main() {}
