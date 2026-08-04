// vybe-test: dart/map_entry_algorithms/entries_map_join_with_custom_separator
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
  var m = {'k1': 1, 'k2': 2};
  var line = m.entries.map((e) => '${e.key}:${e.value}').join('; ');
  __p(line);
}

void main() {
  __vybeMain();
  __check('k1:1; k2:2');
}
