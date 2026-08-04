// vybe-test: dart/map_entry_algorithms/put_if_absent_three_tier_grouping
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
  m.putIfAbsent('g', () => {});
  m['g']!.putIfAbsent('x', () => 0);
  m['g']!.putIfAbsent('y', () => 0);
  m['g']!['x'] = 5;
  __p(m['g']!['x']);
  __p(m['g']!['y']);
}

void main() {
  __vybeMain();
  __check('5\n0');
}
