// vybe-test: dart/sync_star_generators/sync_star_yields_two_values_in_order
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

Iterable<int> pair() sync* { yield 1; yield 2; }
void __vybeMain() {
  var out = <int>[];
  for (var v in pair()) { out.add(v); }
  __p(out.join(','));
}

void main() {
  __vybeMain();
  __check('1,2');
}
