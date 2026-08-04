// vybe-test: dart/spread_collections/set_spread_typed_source
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
  Set<int> src = {5, 6};
  var out = <int>{1, ...src};
  __p(out.contains(6));
  __p(out.length);
}

void main() {
  __vybeMain();
  __check('true\n3');
}
