// vybe-test: dart/sync_star_generators/sync_star_odds_via_continue_skip_evens
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

Iterable<int> odds(int n) sync* {
  for (var i = 0; i < n; i++) {
    if (i % 2 == 0) continue;
    yield i;
  }
}
void __vybeMain() {
  __p(odds(6).join(','));
}

void main() {
  __vybeMain();
  __check('1,3,5');
}
