// vybe-test: dart/sync_star_generators/sync_star_fibonacci_first_ten_sum
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
  var sum = 0;
  for (var v in fib().take(10)) { sum += v; }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('88');
}
