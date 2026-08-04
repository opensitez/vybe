// vybe-test: dart/async_generators/sync_star_basic
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> count() sync* { yield 1; yield 2; yield 3; }

void main() {}
