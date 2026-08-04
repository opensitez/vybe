// vybe-test: dart/async_generators/sync_star_loop
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> range(int end) sync* {
  for (var i = 0; i < end; i++) { yield i; }
}

void main() {}
