// vybe-test: dart/custom_iterator/custom_iterator_indexed_manual_loop
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

class IndexWalk extends IterableBase<int> {
  @override
  Iterator<int> get iterator => IndexWalkIterator();
}
class IndexWalkIterator implements Iterator<int> {
  int _i = -1;
  @override
  int get current => _i;
  @override
  bool moveNext() {
    if (_i < 2) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = IndexWalk().iterator;
  var idx = 0;
  while (it.moveNext()) {
    __p(it.current);
    idx = idx + 1;
  }
  __p(idx);
}

void main() {
  __vybeMain();
  __check('0\n1\n2\n3');
}
