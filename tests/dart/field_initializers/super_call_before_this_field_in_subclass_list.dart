// vybe-test: dart/field_initializers/super_call_before_this_field_in_subclass_list
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

class Parent {
  int p;
  Parent(this.p);
}
class Child extends Parent {
  int c;
  Child(int x, int y) : super(x), c = y + 1;
}
void __vybeMain() {
  __p(Child(5, 10).c);
}

void main() {
  __vybeMain();
  __check('11');
}
