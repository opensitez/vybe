// vybe-test: dart/streams_core/stream_skip_while_skips_matching_prefix
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
  await for (var v in Stream.fromIterable([0, 0, 1, 2, 0]).skipWhile((x) => x == 0)) { out.add(v); }
  __p(out.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('1,2,0');
}
