// vybe-test: dart/unmodifiable_collections/unmod_map_keys_join_preserves_order
// origin: languages/dart/tests/dart/test_unmodifiable_collections.rs

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
  var frozen = Map.unmodifiable({'z': 1, 'a': 2, 'm': 3});
  __p(frozen.keys.join(','));
}

void main() {
  __vybeMain();
  __check('z,a,m');
}
