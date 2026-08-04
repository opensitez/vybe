// vybe-test: dart/spread_collections/set_spread_merges_two_sets
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
  var a = {1, 2};
  var b = {3, 4};
  var merged = {...a, ...b};
  __p(merged.length);
  __p(merged.contains(3));
}

void main() {
  __vybeMain();
  __check('4\ntrue');
}
