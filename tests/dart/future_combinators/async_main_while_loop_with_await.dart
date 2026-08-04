// vybe-test: dart/future_combinators/async_main_while_loop_with_await
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

Future<int> tick(int n) async => n;
Future<void> __vybeMain() async {
  var i = 0;
  var sum = 0;
  while (i < 3) {
    sum = sum + await tick(i + 1);
    i++;
  }
  __p(sum);
}

Future<void> main() async {
  await __vybeMain();
  __check('6');
}
