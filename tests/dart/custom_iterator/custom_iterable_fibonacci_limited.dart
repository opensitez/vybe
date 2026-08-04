// vybe-test: dart/custom_iterator/custom_iterable_fibonacci_limited
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

class Fib extends IterableBase<int> {
  final int count;
  Fib(this.count);
  @override
  Iterator<int> get iterator => FibIterator(count);
}
class FibIterator implements Iterator<int> {
  int _a = 0;
  int _b = 1;
  int _seen = 0;
  final int count;
  FibIterator(this.count);
  @override
  int get current => _a;
  @override
  bool moveNext() {
    if (_seen >= count) {
      return false;
    }
    var next = _a + _b;
    _a = _b;
    _b = next;
    _seen = _seen + 1;
    return true;
  }
}
void __vybeMain() {
  __p(Fib(5).join(','));
}

void main() {
  __vybeMain();
  __check('1,1,2,3,5');
}
