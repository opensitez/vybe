// vybe-test: dart/getters_setters/setter_and_getter_on_same_private_field_independent_instances
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

class Cell {
  int _data = 0;
  int get data {
    return _data;
  }
  set data(int v) {
    _data = v;
  }
}
void __vybeMain() {
  var a = Cell();
  var b = Cell();
  a.data = 3;
  b.data = 7;
  __p(a.data);
  __p(b.data);
}

void main() {
  __vybeMain();
  __check('3\n7');
}
