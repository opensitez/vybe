// vybe-test: dart/async_generators/sync_star_fibonacci
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> fibs() sync* {
  int a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}

void main() {}
