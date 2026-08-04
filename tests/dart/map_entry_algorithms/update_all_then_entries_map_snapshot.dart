// vybe-test: dart/map_entry_algorithms/update_all_then_entries_map_snapshot
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
  var m = {'a': 1, 'b': 2};
  m.updateAll((k, v) => v + 100);
  var snapshot = m.entries.map((e) => e.value).toList()..sort();
  __p(snapshot.join(','));
}

void main() {
  __vybeMain();
  __check('101,102');
}
