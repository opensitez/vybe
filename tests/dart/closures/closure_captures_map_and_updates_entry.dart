// vybe-test: dart/closures/closure_captures_map_and_updates_entry
// origin: languages/dart/tests/dart/test_closures.rs

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
  var counts = <String, int>{'a': 1};
  var bump = (String key) {
    counts[key] = (counts[key] ?? 0) + 1;
  };
  bump('a');
  bump('b');
  __p(counts['a']);
  __p(counts['b']);
}

void main() {
  __vybeMain();
  __check('2\n1');
}
