// vybe-test: dart/streams_core/stream_expand_flattens_inner_lists
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
  var out = <int>[];
  var nested = Stream.fromIterable([[1, 2], [3], [4, 5]]);
  await for (var v in nested.expand((part) => part)) { out.add(v); }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,3,4,5');
}
