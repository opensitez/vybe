// vybe-test: dart/list_core/list_as_map_pairs_indices_with_values
// origin: languages/dart/tests/dart/test_list_core.rs

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
  var list = ['a', 'b', 'c'];
  var map = list.asMap();
  __p(map[1]);
}

void main() {
  __vybeMain();
  __check('b');
}
