// vybe-test: dart/map_entry_algorithms/map_update_with_if_absent_inserts_default
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
  m.update('new', (v) => v + 10, ifAbsent: () => 0);
  __p(m['new']);
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('0\n1');
}
