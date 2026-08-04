// vybe-test: dart/map_core/map_add_all_merges_entries_from_other_map
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
  var m = {'a': 1};
  m.addAll({'b': 2, 'c': 3});
  __p(m.length);
  __p(m['b']);
  __p(m['c']);
}

void main() {
  __vybeMain();
  __check('3\n2\n3');
}
