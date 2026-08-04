// vybe-test: dart/custom_iterator/custom_iterable_subclass_extends_base
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

class BaseRange extends IterableBase<int> {
  final int n;
  BaseRange(this.n);
  @override
  Iterator<int> get iterator => BaseRangeIterator(n);
}
class BaseRangeIterator implements Iterator<int> {
  int _v = 0;
  final int n;
  BaseRangeIterator(this.n);
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < n) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
class DoubleRange extends BaseRange {
  DoubleRange(int n) : super(n);
}
void __vybeMain() {
  __p(DoubleRange(3).join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
