// vybe-test: dart/async_generators/async_star_yield_star
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> first() async* { yield 1; yield 2; }
Stream<int> combined() async* { yield* first(); yield 3; }

void main() {}
