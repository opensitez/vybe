// vybe-test: dart/errors_advanced/on_multiple
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() {
  try {
    throw RangeError('out of range');
  } on FormatException {
    print('format');
  } on RangeError catch (e) {
    print('range');
  }
}