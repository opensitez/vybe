// vybe-test: dart/async_star_generators/async_star_fibonacci_take_seven
// origin: languages/dart/tests/dart/test_async_star_generators.rs

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

Stream<int> fib() async* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
Future<void> __vybeMain() async {
  __p(await fib().take(7).join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('0,1,1,2,3,5,8');
}
