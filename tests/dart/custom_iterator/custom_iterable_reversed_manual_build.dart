// vybe-test: dart/custom_iterator/custom_iterable_reversed_manual_build
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

class Rev extends IterableBase<int> {
  final List<int> data;
  Rev(this.data);
  @override
  Iterator<int> get iterator => RevIterator(data);
}
class RevIterator implements Iterator<int> {
  int _i;
  final List<int> data;
  RevIterator(List<int> data) : data = data, _i = data.length;
  @override
  int get current => data[_i];
  @override
  bool moveNext() {
    if (_i > 0) {
      _i = _i - 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Rev([1, 2, 3]).join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
