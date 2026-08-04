// vybe-test: dart/generics_core/generic_queue_peek_and_dequeue
// origin: languages/dart/tests/dart/test_generics_core.rs

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

class Queue<T> {
  List<T> _data = [];
  void enqueue(T item) {
    _data.add(item);
  }
  T dequeue() {
    return _data.removeAt(0);
  }
  T peek() {
    return _data.first;
  }
}
void __vybeMain() {
  var q = Queue<String>();
  q.enqueue('first');
  q.enqueue('second');
  __p(q.peek());
  __p(q.dequeue());
}

void main() {
  __vybeMain();
  __check('first\nfirst');
}
