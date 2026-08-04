// vybe-test: dart/generic_inference/infer_generic_queue_enqueue_dequeue
// origin: languages/dart/tests/dart/test_generic_inference.rs

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
}
void __vybeMain() {
  var q = Queue();
  q.enqueue('first');
  q.enqueue('second');
  __p(q.dequeue());
}

void main() {
  __vybeMain();
  __check('first');
}
