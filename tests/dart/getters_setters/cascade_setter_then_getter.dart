// vybe-test: dart/getters_setters/cascade_setter_then_getter
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

class Widget {
  int _width = 0;
  int get width {
    return _width;
  }
  set width(int v) {
    _width = v;
  }
}
void __vybeMain() {
  var w = Widget();
  w..width = 20..width = 30;
  __p(w.width);
}

void main() {
  __vybeMain();
  __check('30');
}
