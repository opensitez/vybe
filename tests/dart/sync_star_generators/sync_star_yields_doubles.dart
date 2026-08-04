// vybe-test: dart/sync_star_generators/sync_star_yields_doubles
// origin: languages/dart/tests/dart/test_sync_star_generators.rs

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

Iterable<double> halves() sync* { yield 0.5; yield 1.5; }
void __vybeMain() {
  var sum = 0.0;
  for (var v in halves()) { sum += v; }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('2');
}
