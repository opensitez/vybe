// vybe-test: dart/map_entry_algorithms/map_update_if_absent_not_called_when_key_exists
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
  var m = {'x': 7};
  m.update('x', (v) => v * 2, ifAbsent: () => 100);
  __p(m['x']);
}

void main() {
  __vybeMain();
  __check('14');
}
