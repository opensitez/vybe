// vybe-test: dart/spread_collections/map_spread_empty_into_empty_then_add
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
  var e1 = <String, int>{};
  var e2 = <String, int>{};
  var out = {...e1, ...e2, 'x': 1};
  __p(out.length);
}

void main() {
  __vybeMain();
  __check('1');
}
