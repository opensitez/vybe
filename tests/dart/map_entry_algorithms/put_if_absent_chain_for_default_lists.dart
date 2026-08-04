// vybe-test: dart/map_entry_algorithms/put_if_absent_chain_for_default_lists
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
  var m = <String, List<int>>{};
  m.putIfAbsent('nums', () => []).add(1);
  m.putIfAbsent('nums', () => []).add(2);
  __p(m['nums']!.join(','));
  __p(m['nums']!.length);
}

void main() {
  __vybeMain();
  __check('1,2\n2');
}
