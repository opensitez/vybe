// vybe-test: dart/custom_iterator/custom_iterator_empty_never_moves
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

class EmptyIterator implements Iterator<int> {
  @override
  int get current => 0;
  @override
  bool moveNext() => false;
}
void __vybeMain() {
  var it = EmptyIterator();
  __p(it.moveNext());
}

void main() {
  __vybeMain();
  __check('false');
}
