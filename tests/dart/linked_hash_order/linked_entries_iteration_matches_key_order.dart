// vybe-test: dart/linked_hash_order/linked_entries_iteration_matches_key_order
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
  var m = {'w': 1, 'x': 2, 'y': 3};
  var order = <String>[];
  for (var e in m.entries) { order.add(e.key); }
  __p(order.join(','));
}

void main() {
  __vybeMain();
  __check('w,x,y');
}
