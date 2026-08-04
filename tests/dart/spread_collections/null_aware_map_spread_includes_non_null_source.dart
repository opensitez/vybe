// vybe-test: dart/spread_collections/null_aware_map_spread_includes_non_null_source
// origin: languages/dart/tests/dart/test_spread_collections.rs

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
  Map<String, int>? extra = {'b': 2};
  var out = {'a': 1, ...?extra};
  __p(out.length);
  __p(out['b']);
}

void main() {
  __vybeMain();
  __check('2\n2');
}
