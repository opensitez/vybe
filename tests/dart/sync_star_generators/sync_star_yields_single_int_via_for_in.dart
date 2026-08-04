// vybe-test: dart/sync_star_generators/sync_star_yields_single_int_via_for_in
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

Iterable<int> one() sync* { yield 42; }
void __vybeMain() {
  for (var v in one()) { __p(v); }
}

void main() {
  __vybeMain();
  __check('42');
}
