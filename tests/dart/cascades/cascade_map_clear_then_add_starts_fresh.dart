// vybe-test: dart/cascades/cascade_map_clear_then_add_starts_fresh
// origin: languages/dart/tests/dart/test_cascades.rs

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

void __vybeMain() {
  var scores = {'x': 1, 'y': 2};
  scores..clear()..add('only', 42);
  __p(scores.length);
  __p(scores['only']);
}

void main() {
  __vybeMain();
  __check('1\n42');
}
