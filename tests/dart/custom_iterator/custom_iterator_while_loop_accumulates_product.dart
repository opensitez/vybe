// vybe-test: dart/custom_iterator/custom_iterator_while_loop_accumulates_product
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

class MultIterator implements Iterator<int> {
  int _v = 0;
  final int limit;
  MultIterator(this.limit);
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < limit) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = MultIterator(3);
  var product = 1;
  while (it.moveNext()) {
    product = product * it.current;
  }
  __p(product);
}

void main() {
  __vybeMain();
  __check('6');
}
