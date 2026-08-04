// vybe-test: dart/linked_hash_order/linked_numeric_string_keys_maintain_insertion
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
  var m = <String, int>{};
  m['10'] = 10;
  m['2'] = 2;
  m['1'] = 1;
  __p(m.keys.join(','));
}

void main() {
  __vybeMain();
  __check('10,2,1');
}
