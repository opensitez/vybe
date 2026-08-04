// vybe-test: dart/streams_core/stream_listen_on_done_increments_counter
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

class DoneTracker {
  var doneCount = 0;
}
Future<void> __vybeMain() async {
  var t = DoneTracker();
  var sub = Stream.fromIterable([1]).listen(
    (_) {},
    onDone: () => t.doneCount++,
  );
  await sub.asFuture();
  __p(t.doneCount);
}

Future<void> main() async {
  await __vybeMain();
  __check('1');
}
