// vybe-test: dart/async_generators/async_star_basic
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> count() async* { yield 1; yield 2; yield 3; }

void main() {}
