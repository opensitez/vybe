// vybe-test: dart/custom_iterator/custom_iterable_is_empty_false
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

class OneItem extends IterableBase<int> {
  @override
  Iterator<int> get iterator => OneItemIterator();
}
class OneItemIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 1;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(OneItem().isEmpty);
}

void main() {
  __vybeMain();
  __check('false');
}
