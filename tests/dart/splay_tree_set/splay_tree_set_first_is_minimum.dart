// vybe-test: dart/splay_tree_set/splay_tree_set_first_is_minimum
// origin: languages/dart/tests/dart/test_splay_tree_set.rs

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
  var s = SplayTreeSet<int>();
  s.add(8);
  s.add(2);
  s.add(5);
  __p(s.first);
}

void main() {
  __vybeMain();
  __check('2');
}
