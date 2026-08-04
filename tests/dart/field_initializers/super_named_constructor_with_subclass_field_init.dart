// vybe-test: dart/field_initializers/super_named_constructor_with_subclass_field_init
// origin: languages/dart/tests/dart/test_field_initializers.rs

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
  Base.zero() : n = 0;
}
class Sub extends Base {
  int extra;
  Sub.zero(int e) : super.zero(), extra = e;
}
void __vybeMain() {
  var s = Sub.zero(5);
  __p(s.n + s.extra);
}

void main() {
  __vybeMain();
  __check('5');
}
