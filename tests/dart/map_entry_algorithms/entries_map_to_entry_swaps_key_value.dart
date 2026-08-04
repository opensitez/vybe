// vybe-test: dart/map_entry_algorithms/entries_map_to_entry_swaps_key_value
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
  var m = {'one': 1, 'two': 2};
  var swapped = Map.fromEntries(
    m.entries.map((e) => MapEntry('${e.value}', e.key.length)),
  );
  __p(swapped['1']);
  __p(swapped['2']);
}

void main() {
  __vybeMain();
  __check('3\n3');
}
