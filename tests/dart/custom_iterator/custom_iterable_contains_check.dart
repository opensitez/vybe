// vybe-test: dart/custom_iterator/custom_iterable_contains_check
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

class Small extends IterableBase<int> {
  @override
  Iterator<int> get iterator => SmallIterator();
}
class SmallIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Small().contains(2));
}

void main() {
  __vybeMain();
  __check('true');
}
