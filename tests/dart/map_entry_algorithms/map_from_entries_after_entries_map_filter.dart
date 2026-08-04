// vybe-test: dart/map_entry_algorithms/map_from_entries_after_entries_map_filter
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
  var evens = Map.fromEntries(
    m.entries.where((e) => e.value.isEven).map((e) => MapEntry(e.key, e.value * 10)),
  );
  __p(evens.length);
  __p(evens['b']);
  __p(evens['d']);
}

void main() {
  __vybeMain();
  __check('2\n20\n40');
}
