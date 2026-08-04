// vybe-test: dart/map_entry_algorithms/put_if_absent_sequential_keys_builds_map
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
  var m = <String, int>{};
  m.putIfAbsent('a', () => 1);
  m.putIfAbsent('b', () => 2);
  m.putIfAbsent('c', () => 3);
  __p(m.length);
  __p(m.values.fold(0, (s, v) => s + v));
}

void main() {
  __vybeMain();
  __check('3\n6');
}
