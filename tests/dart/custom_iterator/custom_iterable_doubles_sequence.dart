// vybe-test: dart/custom_iterator/custom_iterable_doubles_sequence
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

class Halves extends IterableBase<double> {
  @override
  Iterator<double> get iterator => HalvesIterator();
}
class HalvesIterator implements Iterator<double> {
  int _n = 0;
  @override
  double get current => _n * 0.5;
  @override
  bool moveNext() {
    if (_n < 4) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Halves().map((d) => d.toString()).join(','));
}

void main() {
  __vybeMain();
  __check('0.5,1.0,1.5,2.0');
}
