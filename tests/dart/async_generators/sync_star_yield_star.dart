// vybe-test: dart/async_generators/sync_star_yield_star
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> first() sync* { yield 1; yield 2; }
Iterable<int> all() sync* { yield* first(); yield 3; }

void main() {}
