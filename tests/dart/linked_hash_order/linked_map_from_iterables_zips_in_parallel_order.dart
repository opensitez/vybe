// vybe-test: dart/linked_hash_order/linked_map_from_iterables_zips_in_parallel_order
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
  var m = Map.fromIterables(['z', 'a', 'm'], [1, 2, 3]);
  __p(m.keys.join(','));
  __p(m.values.join(','));
}

void main() {
  __vybeMain();
  __check('z,a,m\n1,2,3');
}
