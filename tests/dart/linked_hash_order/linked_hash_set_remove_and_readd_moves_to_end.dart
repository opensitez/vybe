// vybe-test: dart/linked_hash_order/linked_hash_set_remove_and_readd_moves_to_end
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
  var s = LinkedHashSet<int>();
  s.add(1);
  s.add(2);
  s.add(3);
  s.remove(2);
  s.add(2);
  __p(s.join(','));
}

void main() {
  __vybeMain();
  __check('1,3,2');
}
