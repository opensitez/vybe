// vybe-test: dart/constructors/named_constructor_with_body_computes_values
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

class Size {
  int w;
  int h;
  Size(this.w, this.h);
  Size.square(int side) : w = side, h = side;
}
void __vybeMain() {
  var s = Size.square(5);
  __p(s.w + s.h);
}

void main() {
  __vybeMain();
  __check('10');
}
