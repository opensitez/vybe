// vybe-test: dart/sync_star_generators/sync_star_generator_with_parameters
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

Iterable<int> repeat(int value, int times) sync* {
  for (var i = 0; i < times; i++) { yield value; }
}
void __vybeMain() {
  __p(repeat(9, 3).join(','));
}

void main() {
  __vybeMain();
  __check('9,9,9');
}
