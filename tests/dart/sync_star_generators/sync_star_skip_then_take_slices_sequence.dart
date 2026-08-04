// vybe-test: dart/sync_star_generators/sync_star_skip_then_take_slices_sequence
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

Iterable<int> seq() sync* {
  for (var i = 0; i < 6; i++) { yield i; }
}
void __vybeMain() {
  __p(seq().skip(2).take(2).join(','));
}

void main() {
  __vybeMain();
  __check('2,3');
}
