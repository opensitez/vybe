// vybe-test: dart/custom_iterator/custom_iterable_empty_is_empty
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

class Nothing extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NothingIterator();
}
class NothingIterator implements Iterator<int> {
  @override
  int get current => 0;
  @override
  bool moveNext() => false;
}
void __vybeMain() {
  __p(Nothing().isEmpty);
}

void main() {
  __vybeMain();
  __check('true');
}
