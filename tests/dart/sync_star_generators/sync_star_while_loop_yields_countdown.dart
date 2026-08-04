// vybe-test: dart/sync_star_generators/sync_star_while_loop_yields_countdown
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

Iterable<int> countdown(int n) sync* {
  while (n > 0) { yield n; n--; }
}
void __vybeMain() {
  __p(countdown(3).join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
