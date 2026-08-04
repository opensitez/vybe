// vybe-test: dart/map_entry_algorithms/remove_where_by_key_prefix
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
  var m = {'temp_a': 1, 'keep_b': 2, 'temp_c': 3};
  m.removeWhere((k, v) => k.startsWith('temp_'));
  __p(m.length);
  __p(m.containsKey('keep_b'));
}

void main() {
  __vybeMain();
  __check('1\ntrue');
}
