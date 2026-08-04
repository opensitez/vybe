// vybe-test: dart/sync_star_generators/sync_star_fibonacci_first_six_via_take
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

Iterable<int> fib() sync* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
void __vybeMain() {
  __p(fib().take(6).join(','));
}

void main() {
  __vybeMain();
  __check('0,1,1,2,3,5');
}
