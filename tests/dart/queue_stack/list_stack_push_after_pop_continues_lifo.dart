// vybe-test: dart/queue_stack/list_stack_push_after_pop_continues_lifo
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
  var stack = <int>[];
  stack.add(1);
  stack.removeLast();
  stack.add(2);
  __p(stack.removeLast());
}

void main() {
  __vybeMain();
  __check('2');
}
