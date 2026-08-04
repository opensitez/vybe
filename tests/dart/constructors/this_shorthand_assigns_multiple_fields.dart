// vybe-test: dart/constructors/this_shorthand_assigns_multiple_fields
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

class RGB {
  int r;
  int g;
  int b;
  RGB(this.r, this.g, this.b);
}
void __vybeMain() {
  var c = RGB(1, 2, 3);
  __p(c.r + c.g + c.b);
}

void main() {
  __vybeMain();
  __check('6');
}
