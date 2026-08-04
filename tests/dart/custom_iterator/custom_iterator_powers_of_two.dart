// vybe-test: dart/custom_iterator/custom_iterator_powers_of_two
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

class Powers extends IterableBase<int> {
  @override
  Iterator<int> get iterator => PowersIterator();
}
class PowersIterator implements Iterator<int> {
  final values = [1, 2, 4, 8];
  int _i = -1;
  @override
  int get current => values[_i];
  @override
  bool moveNext() {
    if (_i + 1 < values.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Powers().join(','));
}

void main() {
  __vybeMain();
  __check('1,2,4,8');
}
