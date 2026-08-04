// vybe-test: dart/async_star_generators/async_star_element_at_reads_position
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

Stream<String> gen() async* { yield 'p'; yield 'q'; yield 'r'; }
Future<void> __vybeMain() async {
  __p(await gen().elementAt(1));
}

Future<void> main() async {
  await __vybeMain();
  __check('q');
}
