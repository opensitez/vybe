// vybe-test: dart/custom_iterator/custom_iterable_join_elements
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

class Letters extends IterableBase<String> {
  @override
  Iterator<String> get iterator => LetterIterator();
}
class LetterIterator implements Iterator<String> {
  int _i = -1;
  final codes = ['x', 'y', 'z'];
  @override
  String get current => codes[_i];
  @override
  bool moveNext() {
    if (_i + 1 < codes.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Letters().join('-'));
}

void main() {
  __vybeMain();
  __check('x-y-z');
}
