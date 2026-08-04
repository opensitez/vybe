// vybe-test: dart/map_entry_algorithms/map_update_all_uses_key_in_computation
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
  var m = {'a': 0, 'b': 0};
  m.updateAll((k, v) => k.length);
  __p(m['a']);
  __p(m['b']);
}

void main() {
  __vybeMain();
  __check('1\n1');
}
