// vybe-test: dart/getters_setters/override_setter_updates_subclass_storage
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

class Base {
  int _v = 0;
  int get v {
    return _v;
  }
  set v(int x) {
    _v = x;
  }
}
class Sub extends Base {
  set v(int x) {
    _v = x * 2;
  }
}
void __vybeMain() {
  var s = Sub();
  s.v = 5;
  __p(s.v);
}

void main() {
  __vybeMain();
  __check('10');
}
