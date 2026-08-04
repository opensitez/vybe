// vybe-test: dart/map_entry_algorithms/entries_fold_concatenates_keys
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
  var m = {'one': 1, 'two': 2, 'three': 3};
  var keys = m.entries.fold('', (acc, e) => acc + e.key);
  __p(keys);
}

void main() {
  __vybeMain();
  __check('onetwothree');
}
