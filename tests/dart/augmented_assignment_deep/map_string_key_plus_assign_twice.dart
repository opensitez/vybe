// vybe-test: dart/augmented_assignment_deep/map_string_key_plus_assign_twice
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
  var m = {'msg': 'a'};
  m['msg'] += 'b';
  m['msg'] += 'c';
  __p(m['msg']);
}

void main() {
  __vybeMain();
  __check('abc');
}
