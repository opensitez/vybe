// vybe-test: dart/custom_iterator/custom_iterable_for_in_loop
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

class RangeIterable extends IterableBase<int> {
  final int start;
  final int end;
  RangeIterable(this.start, this.end);
  @override
  Iterator<int> get iterator => RangeIterator(start, end);
}
class RangeIterator implements Iterator<int> {
  int _current;
  final int end;
  RangeIterator(int start, int end) : _current = start - 1, end = end;
  @override
  int get current => _current;
  @override
  bool moveNext() {
    if (_current < end) {
      _current = _current + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var sum = 0;
  for (var n in RangeIterable(1, 4)) {
    sum = sum + n;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('10');
}
