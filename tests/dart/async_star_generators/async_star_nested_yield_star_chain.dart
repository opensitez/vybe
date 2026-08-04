// vybe-test: dart/async_star_generators/async_star_nested_yield_star_chain
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

Stream<int> a() async* { yield 1; }
Stream<int> b() async* { yield* a(); yield 2; }
Stream<int> c() async* { yield* b(); yield 3; }
Future<void> __vybeMain() async {
  __p(await c().join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3');
}
