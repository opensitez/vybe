// vybe-test: dart/streams_core/stream_listen_data_and_done_together
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

class StreamStats {
  var dataCount = 0;
  var done = false;
}
Future<void> __vybeMain() async {
  var s = StreamStats();
  var sub = Stream.fromIterable([1, 2, 3]).listen(
    (_) => s.dataCount++,
    onDone: () => s.done = true,
  );
  await sub.asFuture();
  __p('${s.dataCount}|${s.done}');
}

Future<void> main() async {
  await __vybeMain();
  __check('3|true');
}
