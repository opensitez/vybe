// vybe-test: dart/async_futures_core/await_chained_async_calls
// origin: languages/dart/tests/dart/test_async_futures_core.rs

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

Future<int> step1() async {
  return 2;
}
Future<int> step2(int n) async {
  return n + 3;
}
void __vybeMain() async {
  var a = await step1();
  var b = await step2(a);
  __p(b);
}

Future<void> main() async {
  await __vybeMain();
  __check('5');
}
