// vybe-test: dart/errors_advanced/catch_with_stack_trace
// origin: languages/dart/tests/dart/test_errors_advanced.rs

void main() {
  try {
    throw Exception('boom');
  } catch (e, s) {
    print('caught');
  }
}