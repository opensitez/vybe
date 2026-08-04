// vybe-test: dart/map_entry_algorithms/put_if_absent_chain_builds_nested_counter
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
  var m = <String, Map<String, int>>{};
  m.putIfAbsent('users', () => {})..putIfAbsent('alice', () => 0);
  m['users']!['alice'] = m['users']!['alice']! + 1;
  __p(m['users']!['alice']);
}

void main() {
  __vybeMain();
  __check('1');
}
