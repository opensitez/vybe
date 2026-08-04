// vybe-test: dart/async_generators/future_error
// origin: languages/dart/tests/dart/test_async_generators.rs

var f = Future<int>.error(Exception('fail'));

void main() {}
