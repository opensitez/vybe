// vybe-test: dart/custom_iterator/custom_iterable_modulo_pattern
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

class ModThree extends IterableBase<int> {
  @override
  Iterator<int> get iterator => ModThreeIterator();
}
class ModThreeIterator implements Iterator<int> {
  int _n = 0;
  @override
  int get current => _n % 3;
  @override
  bool moveNext() {
    if (_n < 6) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(ModThree().join(','));
}

void main() {
  __vybeMain();
  __check('1,2,0,1,2,0');
}
