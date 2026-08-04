// vybe-test: dart/linked_hash_order/linked_spread_overwrite_keeps_first_key_position
// origin: languages/dart/tests/dart/test_linked_hash_order.rs

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
  var a = {'k': 1, 'm': 2};
  var b = {'k': 99, 'n': 3};
  var m = {...a, ...b};
  __p(m.keys.join(','));
  __p(m['k']);
}

void main() {
  __vybeMain();
  __check('k,m,n\n99');
}
