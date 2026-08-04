// vybe-test: dart/async_generators/stream_to_list
// origin: languages/dart/tests/dart/test_async_generators.rs

Future<List<int>> f() async => Stream.fromIterable([1, 2, 3]).toList();

void main() {}
