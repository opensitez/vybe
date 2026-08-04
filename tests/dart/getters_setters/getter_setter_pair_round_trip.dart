// vybe-test: dart/getters_setters/getter_setter_pair_round_trip
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class Box {
  int _size = 1;
  int get size {
    return _size;
  }
  set size(int v) {
    _size = v;
  }
}
void __vybeMain() {
  var b = Box();
  b.size = 10;
  __p(b.size);
}

void main() {
  __vybeMain();
  __check('10');
}
