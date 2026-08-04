// vybe-test: dart/sync_star_generators/sync_star_multiple_yield_star_in_sequence
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

Iterable<int> a() sync* { yield 1; }
Iterable<int> b() sync* { yield 2; }
Iterable<int> both() sync* { yield* a(); yield* b(); yield 3; }
void __vybeMain() {
  __p(both().join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
