// vybe-test: dart/stopwatch_elapsed/stopwatch_lap_compare_before_and_after_stop_same
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
  var ms1 = sw.elapsedMilliseconds;
  var ms2 = sw.elapsedMilliseconds;
  var us1 = sw.elapsedMicroseconds;
  var us2 = sw.elapsedMicroseconds;
  __p(ms1 == ms2);
  __p(us1 == us2);
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
