// vybe-test: dart/mixins_core/mixin_setter_mutates_state
// origin: languages/dart/tests/dart/test_mixins_core.rs

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

mixin Mutable {
  int _v = 0;
  int get v {
    return _v;
  }
  set v(int n) {
    _v = n;
  }
}
class Holder with Mutable {}
void __vybeMain() {
  var h = Holder();
  h.v = 9;
  __p(h.v);
}

void main() {
  __vybeMain();
  __check('9');
}
