// vybe-test: dart/map_entry_algorithms/entries_where_then_map_to_values
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
  var m = {'a': 1, 'b': 10, 'c': 3, 'd': 12};
  var picked = m.entries
      .where((e) => e.value > 5)
      .map((e) => e.value)
      .toList()
    ..sort();
  __p(picked.join(','));
}

void main() {
  __vybeMain();
  __check('10,12');
}
