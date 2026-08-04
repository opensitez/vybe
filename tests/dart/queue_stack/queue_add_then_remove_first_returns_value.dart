// vybe-test: dart/queue_stack/queue_add_then_remove_first_returns_value
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

import 'dart:collection';
void __vybeMain() {
  var q = Queue<int>();
  q.add(10);
  __p(q.removeFirst());
}

void main() {
  __vybeMain();
  __check('10');
}
