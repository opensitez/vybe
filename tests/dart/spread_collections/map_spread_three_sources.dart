// vybe-test: dart/spread_collections/map_spread_three_sources
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
  var a = {'a': 1};
  var b = {'b': 2};
  var c = {'c': 3};
  var all = {...a, ...b, ...c};
  __p(all.length);
}

void main() {
  __vybeMain();
  __check('3');
}
