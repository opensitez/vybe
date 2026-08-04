// vybe-test: dart/custom_iterator/custom_iterable_length_property
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

class Fixed extends IterableBase<int> {
  final int count;
  Fixed(this.count);
  @override
  Iterator<int> get iterator => FixedIterator(count);
  @override
  int get length => count;
}
class FixedIterator implements Iterator<int> {
  int _n = 0;
  final int count;
  FixedIterator(this.count);
  @override
  int get current => _n;
  @override
  bool moveNext() {
    if (_n < count) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Fixed(4).length);
}

void main() {
  __vybeMain();
  __check('4');
}
