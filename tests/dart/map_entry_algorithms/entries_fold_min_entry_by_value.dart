// vybe-test: dart/map_entry_algorithms/entries_fold_min_entry_by_value
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
  var m = {'a': 30, 'b': 10, 'c': 20};
  var minKey = m.entries.fold('?', (best, e) => e.value < m[best]! ? e.key : best);
  __p(minKey);
  __p(m[minKey]);
}

void main() {
  __vybeMain();
  __check('b\n10');
}
