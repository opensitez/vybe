// vybe-test: dart/map_entry_algorithms/entries_map_to_key_value_strings
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
  var m = {'x': 10, 'y': 20};
  var labels = m.entries.map((e) => '${e.key}=${e.value}').toList();
  __p(labels.join('|'));
}

void main() {
  __vybeMain();
  __check('x=10|y=20');
}
