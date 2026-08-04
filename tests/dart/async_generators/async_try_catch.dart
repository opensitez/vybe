// vybe-test: dart/async_generators/async_try_catch
// origin: languages/dart/tests/dart/test_async_generators.rs

Future<void> risky() async { throw Exception('async error'); }
Future<void> main() async {
  try { await risky(); } catch (e) { print('caught'); }
}