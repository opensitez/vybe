// vybe-test: dart/map_entry_algorithms/put_if_absent_chain_on_same_key_returns_existing
// origin: languages/dart/tests/dart/test_map_entry_algorithms.rs

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
  var a = m.putIfAbsent('k', () => 10);
  var b = m.putIfAbsent('k', () => 99);
  __p(a);
  __p(b);
  __p(m['k']);
}

void main() {
  __vybeMain();
  __check('10\n10\n10');
}
