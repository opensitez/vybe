// vybe-test: dart/classes_core/super_used_in_constructor_initializer_with_body
// origin: languages/dart/tests/dart/test_classes_core.rs

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
class Sub extends Base {
  int m;
  Sub(int a, int b) : super(a) {
    m = b;
  }
}
void __vybeMain() {
  var s = Sub(2, 3);
  __p(s.n + s.m);
}

void main() {
  __vybeMain();
  __check('5');
}
