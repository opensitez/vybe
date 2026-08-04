// vybe-test: dart/async_generators/stream_where_method
// origin: languages/dart/tests/dart/test_async_generators.rs

var s = Stream.fromIterable([1, 2, 3, 4]).where((x) => x > 2);

void main() {}
