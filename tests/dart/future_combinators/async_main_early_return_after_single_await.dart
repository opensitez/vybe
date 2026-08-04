// vybe-test: dart/future_combinators/async_main_early_return_after_single_await
// origin: languages/dart/tests/dart/test_future_combinators.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

Future<int> load() async => 50;
Future<int> run() async {
  var v = await load();
  return v;
}
Future<void> __vybeMain() async {
  __p(await run());
}

Future<void> main() async {
  await __vybeMain();
  __check('50');
}
