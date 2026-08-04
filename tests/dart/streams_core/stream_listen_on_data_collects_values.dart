// vybe-test: dart/streams_core/stream_listen_on_data_collects_values
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

class DataSink {
  var values = <int>[];
}
Future<void> __vybeMain() async {
  var sink = DataSink();
  var sub = Stream.fromIterable([4, 5, 6]).listen((v) => sink.values.add(v));
  await sub.asFuture();
  __p(sink.values.join(','));
}

Future<void> main() async {
  await __vybeMain();
  __check('4,5,6');
}
