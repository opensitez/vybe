// vybe-test: dart/async_star_generators/async_star_yield_star_from_sync_iterable
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

Stream<int> fromList() async* { yield* [3, 4, 5]; }
Future<void> __vybeMain() async {
  __p(await fromList().join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('3,4,5');
}
