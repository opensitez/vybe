// vybe-test: dart/streams_core/stream_contains_finds_present_value
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
  var found = await Stream.fromIterable([2, 4, 6]).contains(4);
  __p(found);
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
