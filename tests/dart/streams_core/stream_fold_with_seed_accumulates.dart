// vybe-test: dart/streams_core/stream_fold_with_seed_accumulates
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
  var total = await Stream.fromIterable([1, 2, 3]).fold(10, (acc, v) => acc + v);
  __p(total);
}

Future<void> main() async {
  await __vybeMain();
  __check('16');
}
