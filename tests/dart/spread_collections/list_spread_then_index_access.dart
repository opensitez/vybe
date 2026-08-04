// vybe-test: dart/spread_collections/list_spread_then_index_access
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
  var src = [5, 6];
  var dst = [1, ...src, 7];
  __p(dst[1]);
  __p(dst[2]);
}

void main() {
  __vybeMain();
  __check('5\n6');
}
