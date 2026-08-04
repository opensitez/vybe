// vybe-test: dart/async_generators/yield_star
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> a() sync* { yield 1; } Iterable<int> b() sync* { yield* a(); yield 2; }

void main() {}
