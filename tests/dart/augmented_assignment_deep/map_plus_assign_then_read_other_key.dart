// vybe-test: dart/augmented_assignment_deep/map_plus_assign_then_read_other_key
// origin: languages/dart/tests/dart/test_augmented_assignment_deep.rs

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
  var m = {'a': 1, 'b': 10};
  m['a'] += 2;
  __p(m['b']);
  __p(m['a']);
}

void main() {
  __vybeMain();
  __check('10\n3');
}
