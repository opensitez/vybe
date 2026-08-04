// vybe-test: dart/classes_core/constructor_initializes_multiple_this_fields
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

class Triple {
  int a;
  int b;
  int c;
  Triple(this.a, this.b, this.c);
}
void __vybeMain() {
  var t = Triple(1, 2, 3);
  __p(t.a + t.b + t.c);
}

void main() {
  __vybeMain();
  __check('6');
}
