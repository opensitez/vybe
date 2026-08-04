// vybe-test: dart/async_generators/yield_in_async_star
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> gen() async* { yield 1; yield 2; yield 3; }

void main() {}
