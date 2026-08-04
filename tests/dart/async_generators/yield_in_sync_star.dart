// vybe-test: dart/async_generators/yield_in_sync_star
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> gen() sync* { yield 1; yield 2; yield 3; }

void main() {}
