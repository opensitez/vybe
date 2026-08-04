// vybe-test: dart/cascades/cascade_map_add_overwrites_existing_key_via_chain
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
  var scores = {'k': 5};
  scores..add('k', 9)..add('k', 11);
  __p(scores['k']);
}

void main() {
  __vybeMain();
  __check('11');
}
