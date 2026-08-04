// vybe-test: dart/custom_iterator/custom_iterable_iterator_getter_returns_new_instance
// origin: languages/dart/tests/dart/test_custom_iterator.rs

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

class RepeatOne extends IterableBase<int> {
  @override
  Iterator<int> get iterator => OneIterator();
}
class OneIterator implements Iterator<int> {
  bool _moved = false;
  @override
  int get current => 7;
  @override
  bool moveNext() {
    if (!_moved) {
      _moved = true;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var first = RepeatOne().iterator;
  var second = RepeatOne().iterator;
  first.moveNext();
  second.moveNext();
  __p(first.current + second.current);
}

void main() {
  __vybeMain();
  __check('14');
}
