// vybe-test: dart/cascades/cascade_map_remove_then_add_replaces_entry
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
  var scores = {'old': 1, 'keep': 2};
  scores..remove('old')..add('new', 99);
  __p(scores.containsKey('old'));
  __p(scores['new']);
}

void main() {
  __vybeMain();
  __check('false\n99');
}
