// vybe-test: dart/custom_iterator/custom_iterator_countdown_sequence
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

class Countdown extends IterableBase<int> {
  final int start;
  Countdown(this.start);
  @override
  Iterator<int> get iterator => CountdownIterator(start);
}
class CountdownIterator implements Iterator<int> {
  int _n;
  int _current = 0;
  CountdownIterator(int start) : _n = start;
  @override
  int get current => _current;
  @override
  bool moveNext() {
    if (_n > 0) {
      _current = _n;
      _n = _n - 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var text = '';
  for (var n in Countdown(3)) {
    text = text + n.toString();
  }
  __p(text);
}

void main() {
  __vybeMain();
  __check('321');
}
