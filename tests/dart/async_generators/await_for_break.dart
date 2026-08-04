// vybe-test: dart/async_generators/await_for_break
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<int> nums() async* { for (var i = 0; i < 10; i++) yield i; }
Future<void> main() async {
  await for (var n in nums()) {
    if (n > 3) break;
    print(n);
  }
}