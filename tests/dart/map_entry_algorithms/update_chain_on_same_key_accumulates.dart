// vybe-test: dart/map_entry_algorithms/update_chain_on_same_key_accumulates
// origin: languages/dart/tests/dart/test_map_entry_algorithms.rs

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
  var m = {'score': 0};
  m.update('score', (v) => v + 5);
  m.update('score', (v) => v + 3);
  m.update('score', (v) => v + 2);
  __p(m['score']);
}

void main() {
  __vybeMain();
  __check('10');
}
