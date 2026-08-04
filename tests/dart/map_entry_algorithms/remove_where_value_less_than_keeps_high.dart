// vybe-test: dart/map_entry_algorithms/remove_where_value_less_than_keeps_high
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
  var m = {'w': 100, 'x': 5, 'y': 50, 'z': 1};
  m.removeWhere((k, v) => v < 10);
  __p(m.length);
  __p(m.keys.join(','));
}

void main() {
  __vybeMain();
  __check('2\nw,y');
}
