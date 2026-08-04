// vybe-test: dart/generics/generic_stack
// origin: languages/dart/tests/dart/test_generics.rs

class Stack<T> {
  List<T> _items = [];
  void push(T item) { _items.add(item); }
  T pop() { return _items.removeLast(); }
  bool get isEmpty => _items.isEmpty;
}

void main() {}
