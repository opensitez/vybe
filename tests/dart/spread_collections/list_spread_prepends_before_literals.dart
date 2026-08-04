// vybe-test: dart/spread_collections/list_spread_prepends_before_literals
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
  var tail = [3, 4];
  var all = [1, 2, ...tail];
  __p(all.first);
  __p(all.last);
}

void main() {
  __vybeMain();
  __check('1\n4');
}
