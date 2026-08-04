// vybe-test: dart/async_star_generators/async_star_first_and_last_selectors
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

Stream<int> gen() async* { yield 10; yield 20; yield 30; }
Future<void> __vybeMain() async {
  __p(await gen().first);
  __p(await gen().last);
}

Future<void> main() async {
  await __vybeMain();
  __check('10\n30');
}
