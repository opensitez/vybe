// vybe-test: dart/linked_hash_order/linked_remove_where_then_put_if_absent
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
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.removeWhere((k, v) => k == 'b');
  m.putIfAbsent('d', () => 4);
  __p(m.keys.join(','));
}

void main() {
  __vybeMain();
  __check('a,c,d');
}
