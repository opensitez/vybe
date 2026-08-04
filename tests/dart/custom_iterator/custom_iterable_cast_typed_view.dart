// vybe-test: dart/custom_iterator/custom_iterable_cast_typed_view
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

class AnyNums extends IterableBase<num> {
  @override
  Iterator<num> get iterator => AnyNumsIterator();
}
class AnyNumsIterator implements Iterator<num> {
  int _v = 0;
  @override
  num get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(AnyNums().cast<int>().join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
