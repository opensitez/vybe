// vybe-test: dart/async_generators/future_delayed
// origin: languages/dart/tests/dart/test_async_generators.rs

var f = Future.delayed(Duration(milliseconds: 100), () => 42);

void main() {}
