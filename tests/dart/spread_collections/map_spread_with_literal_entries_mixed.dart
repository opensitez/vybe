// vybe-test: dart/spread_collections/map_spread_with_literal_entries_mixed
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
  var base = {'a': 1};
  var out = {'z': 0, ...base, 'b': 2};
  __p(out['z']);
  __p(out['a']);
  __p(out['b']);
}

void main() {
  __vybeMain();
  __check('0\n1\n2');
}
