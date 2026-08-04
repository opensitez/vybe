// vybe-test: dart/async_generators/async_await_call
// origin: languages/dart/tests/dart/test_async_generators.rs

Future<int> compute() async { return 99; }
Future<void> main() async { var result = await compute(); print(result); }