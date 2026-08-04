// vybe-test: dart/custom_iterator/custom_iterable_expand_flatten
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

class Pairs extends IterableBase<List<int>> {
  @override
  Iterator<List<int>> get iterator => PairsIterator();
}
class PairsIterator implements Iterator<List<int>> {
  int _step = 0;
  @override
  List<int> get current => _step == 0 ? [1, 2] : [3];
  @override
  bool moveNext() {
    if (_step < 2) {
      _step = _step + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Pairs().expand((p) => p).join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
