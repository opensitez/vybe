// vybe-test: dart/map_entry_algorithms/entries_fold_finds_longest_key
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
  var m = {'a': 1, 'bb': 2, 'ccc': 3};
  var longest = m.entries.fold('', (best, e) => e.key.length > best.length ? e.key : best);
  __p(longest);
}

void main() {
  __vybeMain();
  __check('ccc');
}
