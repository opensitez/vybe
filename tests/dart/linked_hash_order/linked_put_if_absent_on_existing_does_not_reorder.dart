// vybe-test: dart/linked_hash_order/linked_put_if_absent_on_existing_does_not_reorder
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
  var m = {'a': 1, 'b': 2};
  m.putIfAbsent('a', () => 99);
  __p(m.keys.join(','));
  __p(m['a']);
}

void main() {
  __vybeMain();
  __check('a,b\n1');
}
