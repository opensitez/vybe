// vybe-test: dart/cascades/cascade_map_update_chain_mutates_values_in_place
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
  var scores = {'n': 5};
  scores..update('n', (v) => v + 1)..update('n', (v) => v * 2);
  __p(scores['n']);
}

void main() {
  __vybeMain();
  __check('12');
}
