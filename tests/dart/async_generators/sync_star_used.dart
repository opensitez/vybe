// vybe-test: dart/async_generators/sync_star_used
// origin: languages/dart/tests/dart/test_async_generators.rs

Iterable<int> evens(int n) sync* { for (var i = 0; i < n; i += 2) yield i; } void main() { var list = evens(10).toList(); }