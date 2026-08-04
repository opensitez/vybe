// vybe-test: dart/async_generators/future_microtask
// origin: languages/dart/tests/dart/test_async_generators.rs

var f = Future.microtask(() => 42);

void main() {}
