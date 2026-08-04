// vybe-test: dart/async_star_generators/async_star_continue_skips_even_yields
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

Stream<int> odds(int n) async* {
  for (var i = 0; i < n; i++) {
    if (i % 2 == 0) continue;
    yield i;
  }
}
Future<void> __vybeMain() async {
  __p(await odds(6).join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,3,5');
}
