// vybe-test: dart/custom_iterator/custom_iterable_single_length_one
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

class Single extends IterableBase<int> {
  @override
  Iterator<int> get iterator => SingleIterator();
  @override
  int get length => 1;
}
class SingleIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 9;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Single().length);
}

void main() {
  __vybeMain();
  __check('1');
}
