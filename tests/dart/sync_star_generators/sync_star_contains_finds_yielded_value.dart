// vybe-test: dart/sync_star_generators/sync_star_contains_finds_yielded_value
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

Iterable<int> gen() sync* { yield 11; yield 22; }
void __vybeMain() {
  __p(gen().contains(22));
}

void main() {
  __vybeMain();
  __check('true');
}
