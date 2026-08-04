// vybe-test: dart/splay_tree_map/splay_tree_map_keys_ascending
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
  var m = SplayTreeMap<int, String>();
  m[3] = "c";
  m[1] = "a";
  m[2] = "b";
  for (var k in m.keys) {
    __p(k);
  }
}

void main() {
  __vybeMain();
  __check('1\n2\n3');
}
