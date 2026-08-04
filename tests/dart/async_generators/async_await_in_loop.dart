// vybe-test: dart/async_generators/async_await_in_loop
// origin: languages/dart/tests/dart/test_async_generators.rs

Future<int> fetch(int i) async { return i * 2; }
Future<void> main() async {
  for (var i = 0; i < 3; i++) {
    var v = await fetch(i);
    print(v);
  }
}