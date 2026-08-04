// vybe-test: dart/async_star_generators/async_star_multiple_streams_independent
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

Stream<int> gen() async* { yield 1; yield 2; }
Future<void> __vybeMain() async {
  __p(await gen().first);
  __p(await gen().first);
}

Future<void> main() async {
  await __vybeMain();
  __check('1\n1');
}
