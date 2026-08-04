// vybe-test: dart/stopwatch_elapsed/stopwatch_reset_zeros_elapsed
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
  sw.stop();
  sw.reset();
  __p(sw.elapsed.inMicroseconds);
  __p(sw.elapsedMilliseconds);
}

void main() {
  __vybeMain();
  __check('0\n0');
}
