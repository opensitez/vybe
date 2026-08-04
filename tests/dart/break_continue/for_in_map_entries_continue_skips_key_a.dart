// vybe-test: dart/break_continue/for_in_map_entries_continue_skips_key_a
// origin: languages/dart/tests/dart/test_break_continue.rs

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
  var m = {'a': 1, 'b': 2, 'c': 3};
  var sum = 0;
  for (var e in m.entries) {
    if (e.key == 'a') continue;
    sum += e.value;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('5');
}
