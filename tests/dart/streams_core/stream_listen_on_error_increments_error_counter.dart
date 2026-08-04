// vybe-test: dart/streams_core/stream_listen_on_error_increments_error_counter
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

class ErrorTracker {
  var errorCount = 0;
  String? lastError;
}
Future<void> __vybeMain() async {
  var t = ErrorTracker();
  var sub = Stream<int>.error('boom').listen(
    (_) {},
    onError: (e) {
      t.errorCount++;
      t.lastError = '$e';
    },
  );
  try {
    await sub.asFuture();
  } catch (_) {}
  __p('${t.errorCount}|${t.lastError}');
}

Future<void> main() async {
  await __vybeMain();
  __check('1|boom');
}
