// vybe-test: dart/streams_core/stream_to_list_after_map_transform
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
  var list = await Stream.fromIterable([1, 2, 3]).map((x) => x + 1).toList();
  __p(list.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('2,3,4');
}
