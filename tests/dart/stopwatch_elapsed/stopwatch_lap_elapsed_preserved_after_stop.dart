// vybe-test: dart/stopwatch_elapsed/stopwatch_lap_elapsed_preserved_after_stop
// origin: languages/dart/tests/dart/test_stopwatch_elapsed.rs

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

void __vybeMain() {
  var sw = Stopwatch()..start();
  sw.stop();
  var lapMs = sw.elapsedMilliseconds;
  var lapUs = sw.elapsedMicroseconds;
  __p(lapMs >= 0);
  __p(lapUs >= 0);
  __p(sw.isRunning);
}

void main() {
  __vybeMain();
  __check('true\ntrue\nfalse');
}
