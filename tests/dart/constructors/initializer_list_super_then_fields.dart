// vybe-test: dart/constructors/initializer_list_super_then_fields
// origin: languages/dart/tests/dart/test_constructors.rs

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

class A {
  int x;
  A(this.x);
}
class B extends A {
  int y;
  B(int a, int b) : super(a), y = b;
}
void __vybeMain() {
  var b = B(1, 2);
  __p(b.x + b.y);
}

void main() {
  __vybeMain();
  __check('3');
}
