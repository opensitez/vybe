// vybe-test: dart/queue_stack/list_stack_sublist_copy_does_not_change_stack
// origin: languages/dart/tests/dart/test_queue_stack.rs

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
  var stack = <int>[1, 2, 3];
  var copy = stack.sublist(0);
  copy.add(4);
  __p(stack.length);
  __p(stack.last);
}

void main() {
  __vybeMain();
  __check('3\n3');
}
