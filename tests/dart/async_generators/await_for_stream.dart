// vybe-test: dart/async_generators/await_for_stream
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> nums() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  await for (var n in nums()) {
    print(n);
  }
}