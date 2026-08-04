// vybe-test: dart/getters_setters/setter_validates_non_negative_value
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

class Gauge {
  int _level = 0;
  int get level {
    return _level;
  }
  set level(int v) {
    if (v < 0) {
      v = 0;
    }
    _level = v;
  }
}
void __vybeMain() {
  var g = Gauge();
  g.level = -5;
  __p(g.level);
}

void main() {
  __vybeMain();
  __check('0');
}
