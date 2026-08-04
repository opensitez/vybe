// vybe-test: dart/mixin_linearization/mixin_on_base_with_subclass_constructor
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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
  int v;
  Base(this.v);
}
mixin Scale on Base {
  int scaled() {
    return v * 2;
  }
}
class Sub extends Base with Scale {
  Sub(int x) : super(x);
}
void __vybeMain() {
  __p(Sub(4).scaled());
}

void main() {
  __vybeMain();
  __check('8');
}
