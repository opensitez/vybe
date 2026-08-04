// vybe-test: dart/mixins_core/mixin_on_with_named_constructor_on_subclass
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

class Base {
  int v;
  Base(this.v);
}
mixin Scale on Base {
  int scaled() {
    return v * 10;
  }
}
class Sub extends Base with Scale {
  Sub.zero() : super(0);
}
void __vybeMain() {
  __p(Sub.zero().scaled());
}

void main() {
  __vybeMain();
  __check('0');
}
