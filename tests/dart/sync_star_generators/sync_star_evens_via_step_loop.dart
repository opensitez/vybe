// vybe-test: dart/sync_star_generators/sync_star_evens_via_step_loop
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

Iterable<int> evens(int limit) sync* {
  for (var i = 0; i < limit; i += 2) { yield i; }
}
void __vybeMain() {
  __p(evens(8).join(','));
}

void main() {
  __vybeMain();
  __check('0,2,4,6');
}
