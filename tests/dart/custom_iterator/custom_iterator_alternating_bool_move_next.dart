// vybe-test: dart/custom_iterator/custom_iterator_alternating_bool_move_next
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

class AltIterator implements Iterator<int> {
  int _count = 0;
  @override
  int get current => _count;
  @override
  bool moveNext() {
    if (_count < 3) {
      _count = _count + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  var it = AltIterator();
  __p(it.moveNext());
  __p(it.moveNext());
  __p(it.moveNext());
  __p(it.moveNext());
}

void main() {
  __vybeMain();
  __check('true\ntrue\ntrue\nfalse');
}
