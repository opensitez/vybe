// vybe-test: dart/async_star_generators/async_star_while_loop_yields_countdown
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

Stream<int> down(int n) async* {
  while (n > 0) { yield n; n--; }
}
Future<void> __vybeMain() async {
  __p(await down(3).join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('3,2,1');
}
