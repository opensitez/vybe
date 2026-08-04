// vybe-test: dart/custom_iterator/custom_iterator_even_numbers_only
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

class Evens extends IterableBase<int> {
  final int max;
  Evens(this.max);
  @override
  Iterator<int> get iterator => EvensIterator(max);
}
class EvensIterator implements Iterator<int> {
  int _n = 0;
  final int max;
  EvensIterator(this.max);
  @override
  int get current => _n;
  @override
  bool moveNext() {
  _n = _n + 2;
    if (_n <= max) {
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var list = <int>[];
  for (var n in Evens(6)) {
    list.add(n);
  }
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('2,4,6');
}
