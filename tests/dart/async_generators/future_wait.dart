// vybe-test: dart/async_generators/future_wait
// origin: languages/dart/tests/dart/test_async_generators.rs

var f1 = Future.value(1); var f2 = Future.value(2); var all = Future.wait([f1, f2]);

void main() {}
