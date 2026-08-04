// vybe-test: dart/classes_core/this_field_constructor_shorthand
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

class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
}
void __vybeMain() {
  var p = Pair(2, 3);
  __p(p.a + p.b);
}

void main() {
  __vybeMain();
  __check('5');
}
