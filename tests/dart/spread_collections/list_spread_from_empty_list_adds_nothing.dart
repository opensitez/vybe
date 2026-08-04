// vybe-test: dart/spread_collections/list_spread_from_empty_list_adds_nothing
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
  var empty = <int>[];
  var merged = [...empty, 1, 2];
  __p(merged.length);
  __p(merged.join(','));
}

void main() {
  __vybeMain();
  __check('2\n1,2');
}
