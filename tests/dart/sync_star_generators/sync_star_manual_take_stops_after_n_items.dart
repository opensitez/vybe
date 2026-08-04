// vybe-test: dart/sync_star_generators/sync_star_manual_take_stops_after_n_items
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

Iterable<int> naturals() sync* {
  var n = 1;
  while (true) { yield n; n++; }
}
void __vybeMain() {
  var out = <int>[];
  var taken = 0;
  for (var v in naturals()) {
    out.add(v);
    taken++;
    if (taken == 3) break;
  }
  __p(out.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
