// vybe-test: dart/sync_star_generators/sync_star_infinite_generator_safe_with_take
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

Iterable<int> infinite() sync* {
  var n = 0;
  while (true) { yield n; n++; }
}
void __vybeMain() {
  __p(infinite().take(4).join(','));
}

void main() {
  __vybeMain();
  __check('0,1,2,3');
}
