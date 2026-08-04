// vybe-test: dart/unmodifiable_collections/unmod_map_from_mutable_source_captures_snapshot
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
  var src = <String, int>{'a': 1};
  var frozen = Map.unmodifiable(src);
  src['b'] = 2;
  __p(frozen.length);
  __p(frozen.containsKey('b'));
}

void main() {
  __vybeMain();
  __check('1\nfalse');
}
