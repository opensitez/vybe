// vybe-test: dart/future_combinators/async_main_three_awaits_accumulates_sum
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

Future<int> one() async => 1;
Future<int> two() async => 2;
Future<int> three() async => 3;
Future<void> __vybeMain() async {
  var total = await one() + await two() + await three();
  __p(total);
}

Future<void> main() async {
  await __vybeMain();
  __check('6');
}
