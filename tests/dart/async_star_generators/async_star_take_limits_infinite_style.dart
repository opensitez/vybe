// vybe-test: dart/async_star_generators/async_star_take_limits_infinite_style
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

Stream<int> naturals() async* {
  var n = 1;
  while (true) { yield n; n++; }
}
Future<void> __vybeMain() async {
  __p(await naturals().take(4).join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3,4');
}
