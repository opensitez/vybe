// vybe-test: dart/getters_setters/setter_clamps_to_maximum
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

class Volume {
  int _value = 0;
  int get value {
    return _value;
  }
  set value(int v) {
    if (v > 100) {
      v = 100;
    }
    _value = v;
  }
}
void __vybeMain() {
  var v = Volume();
  v.value = 150;
  __p(v.value);
}

void main() {
  __vybeMain();
  __check('100');
}
