// vybe-test: dart/cascades/cascade_map_add_chain_inserts_multiple_keys
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
  var scores = <String, int>{};
  scores..add('alice', 10)..add('bob', 20)..add('carol', 30);
  __p(scores['alice']);
  __p(scores['bob']);
  __p(scores['carol']);
}

void main() {
  __vybeMain();
  __check('10\n20\n30');
}
