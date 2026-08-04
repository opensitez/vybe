// vybe-test: dart/super_parameters/super_param_three_positional_forwards
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Triple {
  int a;
  int b;
  int c;
  Triple(this.a, this.b, this.c);
}
class Child extends Triple {
  Child(super.a, super.b, super.c);
}
void __vybeMain() {
  __p(Child(1, 2, 3).b);
}

void main() {
  __vybeMain();
  __check('2');
}
