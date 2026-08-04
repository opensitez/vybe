// vybe-test: dart/linked_hash_order/linked_literal_with_duplicate_keys_keeps_last_value
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
  var m = {'k': 1, 'k': 2};
  __p(m.keys.join(','));
  __p(m['k']);
  __p(m.length);
}

void main() {
  __vybeMain();
  __check('k\n2\n1');
}
