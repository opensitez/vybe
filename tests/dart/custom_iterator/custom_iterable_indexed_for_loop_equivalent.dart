// vybe-test: dart/custom_iterator/custom_iterable_indexed_for_loop_equivalent
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

class ZeroToTwo extends IterableBase<int> {
  @override
  Iterator<int> get iterator => ZeroToTwoIterator();
}
class ZeroToTwoIterator implements Iterator<int> {
  int _v = -1;
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
  var total = 0;
  for (var i = 0; i < 3; i++) {
    total = total + i;
  }
  var iterTotal = 0;
  for (var n in ZeroToTwo()) {
    iterTotal = iterTotal + n;
  }
  __p(total);
  __p(iterTotal);
}

void main() {
  __vybeMain();
  __check('3\n3');
}
