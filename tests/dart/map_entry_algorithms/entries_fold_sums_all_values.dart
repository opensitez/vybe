// vybe-test: dart/map_entry_algorithms/entries_fold_sums_all_values
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
  var m = {'a': 10, 'b': 20, 'c': 30};
  var total = m.entries.fold(0, (sum, e) => sum + e.value);
  __p(total);
}

void main() {
  __vybeMain();
  __check('60');
}
