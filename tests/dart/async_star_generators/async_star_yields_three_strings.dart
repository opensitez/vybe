// vybe-test: dart/async_star_generators/async_star_yields_three_strings
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

Stream<String> tags() async* { yield 'x'; yield 'y'; yield 'z'; }
Future<void> __vybeMain() async {
  __p(await tags().join(''));
}

Future<void> main() async {
  await __vybeMain();
  __check('xyz');
}
