// vybe-test: dart/map_core/map_values_to_list_matches_insertion_order
// origin: languages/dart/tests/dart/test_map_core.rs

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
  var m = {'a': 10, 'b': 20, 'c': 30};
  var vals = m.values.toList();
  __p(vals.join(','));
}

void main() {
  __vybeMain();
  __check('10,20,30');
}
