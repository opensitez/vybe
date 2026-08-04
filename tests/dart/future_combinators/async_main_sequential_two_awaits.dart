// vybe-test: dart/future_combinators/async_main_sequential_two_awaits
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

Future<int> stepA() async => 10;
Future<int> stepB(int n) async => n + 5;
Future<void> __vybeMain() async {
  var a = await stepA();
  var b = await stepB(a);
  __p(b);
}

Future<void> main() async {
  await __vybeMain();
  __check('15');
}
