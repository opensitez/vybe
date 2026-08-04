// vybe-test: dart/linked_hash_order/linked_hash_map_from_empty_then_sequential_puts
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
  var m = LinkedHashMap<int, String>();
  m[3] = 'three';
  m[1] = 'one';
  m[2] = 'two';
  __p(m.keys.join(','));
  __p(m.values.join(','));
}

void main() {
  __vybeMain();
  __check('3,1,2\nthree,one,two');
}
