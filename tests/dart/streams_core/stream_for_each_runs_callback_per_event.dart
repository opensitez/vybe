// vybe-test: dart/streams_core/stream_for_each_runs_callback_per_event
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
  var sum = 0;
  await Stream.fromIterable([1, 2, 3]).forEach((v) => sum = sum + v);
  __p(sum);
}

Future<void> main() async {
  await __vybeMain();
  __check('6');
}
