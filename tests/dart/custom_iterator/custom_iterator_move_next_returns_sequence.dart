// vybe-test: dart/custom_iterator/custom_iterator_move_next_returns_sequence
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

class ThreeIterator implements Iterator<int> {
  int _step = 0;
  @override
  int get current => _step;
  @override
  bool moveNext() {
    if (_step < 3) {
      _step = _step + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = ThreeIterator();
  var sum = 0;
  while (it.moveNext()) {
    sum = sum + it.current;
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('6');
}
