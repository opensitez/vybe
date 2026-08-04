// vybe-test: dart/custom_iterator/custom_iterator_manual_while_loop_counts
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

class CountIterator implements Iterator<int> {
  int _n = 0;
  final int limit;
  CountIterator(this.limit);
  @override
  int get current => _n;
  @override
  bool moveNext() {
    if (_n < limit) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = CountIterator(5);
  var count = 0;
  while (it.moveNext()) {
    count = count + 1;
  }
  __p(count);
}

void main() {
  __vybeMain();
  __check('5');
}
