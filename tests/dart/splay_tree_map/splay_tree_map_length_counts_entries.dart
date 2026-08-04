// vybe-test: dart/splay_tree_map/splay_tree_map_length_counts_entries
// origin: languages/dart/tests/dart/test_splay_tree_map.rs

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
  var m = SplayTreeMap<int, int>();
  m[10] = 1;
  m[20] = 2;
  m[30] = 3;
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('3');
}
