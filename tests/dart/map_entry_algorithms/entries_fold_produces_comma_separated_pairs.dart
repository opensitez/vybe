// vybe-test: dart/map_entry_algorithms/entries_fold_produces_comma_separated_pairs
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
  var m = {'k1': 1, 'k2': 2};
  var s = m.entries.fold('', (acc, e) {
    if (acc.isEmpty) return '${e.key}:${e.value}';
    return '$acc,${e.key}:${e.value}';
  });
  __p(s.contains('k1:1'));
  __p(s.contains('k2:2'));
}

void main() {
  __vybeMain();
  __check('true\ntrue');
}
