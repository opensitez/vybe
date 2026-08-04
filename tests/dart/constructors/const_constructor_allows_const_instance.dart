// vybe-test: dart/constructors/const_constructor_allows_const_instance
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

class Imm {
  final int x;
  const Imm(this.x);
}
void __vybeMain() {
  const v = Imm(7);
  __p(v.x);
}

void main() {
  __vybeMain();
  __check('7');
}
