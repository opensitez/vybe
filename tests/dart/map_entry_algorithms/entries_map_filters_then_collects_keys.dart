// vybe-test: dart/map_entry_algorithms/entries_map_filters_then_collects_keys
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
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  var bigKeys = m.entries
      .where((e) => e.value > 2)
      .map((e) => e.key)
      .toList();
  __p(bigKeys.join(','));
}

void main() {
  __vybeMain();
  __check('c,d');
}
