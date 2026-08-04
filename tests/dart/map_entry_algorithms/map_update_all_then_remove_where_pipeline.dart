// vybe-test: dart/map_entry_algorithms/map_update_all_then_remove_where_pipeline
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
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.updateAll((k, v) => v * 10);
  m.removeWhere((k, v) => v == 20);
  __p(m.length);
  __p(m['a']);
  __p(m['c']);
}

void main() {
  __vybeMain();
  __check('2\n10\n30');
}
