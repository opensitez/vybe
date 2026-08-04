// vybe-test: dart/streams_core/stream_every_all_match_predicate
// origin: languages/dart/tests/dart/test_streams_core.rs

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

Future<void> __vybeMain() async {
  var ok = await Stream.fromIterable([2, 4, 6]).every((x) => x % 2 == 0);
  __p(ok);
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
