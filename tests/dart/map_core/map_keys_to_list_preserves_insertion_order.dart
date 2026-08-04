// vybe-test: dart/map_core/map_keys_to_list_preserves_insertion_order
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
  var m = {'z': 1, 'a': 2, 'm': 3};
  var keys = m.keys.toList();
  __p(keys.join(','));
}

void main() {
  __vybeMain();
  __check('z,a,m');
}
