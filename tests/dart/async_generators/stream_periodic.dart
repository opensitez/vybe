// vybe-test: dart/async_generators/stream_periodic
// origin: languages/dart/tests/dart/test_async_generators.rs

var s = Stream.periodic(Duration(seconds: 1), (i) => i);

void main() {}
