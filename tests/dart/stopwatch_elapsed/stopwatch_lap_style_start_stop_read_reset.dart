// vybe-test: dart/stopwatch_elapsed/stopwatch_lap_style_start_stop_read_reset
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
  var sw = Stopwatch();
  sw.start();
  __p(sw.isRunning);
  sw.stop();
  var lap = sw.elapsedMilliseconds;
  __p(lap >= 0);
  sw.reset();
  __p(sw.elapsedMilliseconds);
  __p(sw.isRunning);
}

void main() {
  __vybeMain();
  __check('true\ntrue\n0\nfalse');
}
