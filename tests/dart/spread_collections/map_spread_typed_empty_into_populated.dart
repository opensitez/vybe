// vybe-test: dart/spread_collections/map_spread_typed_empty_into_populated
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
  Map<String, int> extra = {'m': 3};
  var out = <String, int>{'n': 4, ...extra};
  __p(out.length);
  __p(out['m']);
}

void main() {
  __vybeMain();
  __check('2\n3');
}
