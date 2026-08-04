// vybe-test: dart/linked_hash_order/linked_update_if_absent_inserts_at_end
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
  var m = {'x': 1};
  m.update('y', (v) => v + 1, ifAbsent: () => 0);
  __p(m.keys.join(','));
  __p(m['y']);
}

void main() {
  __vybeMain();
  __check('x,y\n0');
}
