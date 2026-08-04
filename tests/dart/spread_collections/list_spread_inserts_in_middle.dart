// vybe-test: dart/spread_collections/list_spread_inserts_in_middle
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
  var left = [1, 2];
  var right = [5, 6];
  var mid = [...left, 3, 4, ...right];
  __p(mid.join('-'));
}

void main() {
  __vybeMain();
  __check('1-2-3-4-5-6');
}
