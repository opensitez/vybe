// vybe-test: dart/constructors/generative_positional_constructor_sets_fields
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

class Pair {
  int x;
  int y;
  Pair(this.x, this.y);
}
void __vybeMain() {
  var p = Pair(3, 4);
  __p(p.x + p.y);
}

void main() {
  __vybeMain();
  __check('7');
}
