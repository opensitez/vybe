// vybe-test: dart/generic_inference/infer_generic_stack_push_pop_int
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

class Stack<T> {
  List<T> _items = [];
  void push(T item) {
    _items.add(item);
  }
  T pop() {
    return _items.removeLast();
  }
}
void __vybeMain() {
  var s = Stack();
  s.push(1);
  s.push(2);
  __p(s.pop());
}

void main() {
  __vybeMain();
  __check('2');
}
