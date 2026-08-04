// vybe-test: dart/sync_star_generators/sync_star_expand_flattens_inner_iterables
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

Iterable<List<int>> gen() sync* {
  yield [1, 2];
  yield [3];
}
void __vybeMain() {
  __p(gen().expand((p) => p).join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
