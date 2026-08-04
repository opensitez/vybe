// vybe-test: dart/map_entry_algorithms/entries_map_to_map_with_transformed_values
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
  var m = {'p': 1, 'q': 2};
  var squared = {for (var e in m.entries) e.key: e.value * e.value};
  __p(squared['p']);
  __p(squared['q']);
}

void main() {
  __vybeMain();
  __check('1\n4');
}
