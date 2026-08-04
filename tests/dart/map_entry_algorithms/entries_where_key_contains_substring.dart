// vybe-test: dart/map_entry_algorithms/entries_where_key_contains_substring
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
  var m = {'user_alice': 1, 'admin_bob': 2, 'user_carol': 3};
  var userCount = m.entries.where((e) => e.key.contains('user_')).length;
  __p(userCount);
}

void main() {
  __vybeMain();
  __check('2');
}
