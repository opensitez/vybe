// vybe-test: dart/stopwatch_elapsed/stopwatch_lap_read_while_running_structural
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
  var lap1Running = sw.isRunning;
  var lap1ElapsedOk = sw.elapsed.inMicroseconds >= 0;
  sw.stop();
  var lap1Stopped = sw.isRunning;
  __p(lap1Running);
  __p(lap1ElapsedOk);
  __p(lap1Stopped);
}

void main() {
  __vybeMain();
  __check('true\ntrue\nfalse');
}
