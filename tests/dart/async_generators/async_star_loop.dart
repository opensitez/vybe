// vybe-test: dart/async_generators/async_star_loop
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> range(int n) async* {
  for (var i = 0; i < n; i++) { yield i; }
}

void main() {}
