// vybe-test: dart/getters_setters/setter_updates_dependent_getter
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

class Pair {
  int _a = 1;
  int _b = 2;
  int get a {
    return _a;
  }
  set a(int v) {
    _a = v;
  }
  int get total {
    return _a + _b;
  }
}
void __vybeMain() {
  var p = Pair();
  p.a = 5;
  __p(p.total);
}

void main() {
  __vybeMain();
  __check('7');
}
