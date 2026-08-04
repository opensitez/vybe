// vybe-test: dart/augmented_assignment_deep/list_last_index_plus_assign
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
  var xs = [10, 20, 30];
  xs[2] += 7;
  __p(xs[2]);
}

void main() {
  __vybeMain();
  __check('37');
}
