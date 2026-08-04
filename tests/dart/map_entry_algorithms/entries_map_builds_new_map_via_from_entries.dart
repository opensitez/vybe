// vybe-test: dart/map_entry_algorithms/entries_map_builds_new_map_via_from_entries
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
  var inverted = Map.fromEntries(
    m.entries.map((e) => MapEntry('${e.key}_x', e.value + 10)),
  );
  __p(inverted['a_x']);
  __p(inverted['b_x']);
  __p(inverted.length);
}

void main() {
  __vybeMain();
  __check('11\n12\n2');
}
