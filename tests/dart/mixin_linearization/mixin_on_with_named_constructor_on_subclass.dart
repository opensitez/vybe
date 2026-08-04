// vybe-test: dart/mixin_linearization/mixin_on_with_named_constructor_on_subclass
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
  int n;
  Base(this.n);
}
mixin Double on Base {
  int twice() {
    return n * 2;
  }
}
class Sub extends Base with Double {
  Sub.zero() : super(0);
}
void __vybeMain() {
  __p(Sub.zero().twice());
}

void main() {
  __vybeMain();
  __check('0');
}
