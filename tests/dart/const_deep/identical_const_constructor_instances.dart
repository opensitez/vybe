// vybe-test: dart/const_deep/identical_const_constructor_instances
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Token {
  final int id;
  const Token(this.id);
}
void __vybeMain() {
  __p(identical(const Token(1), const Token(1)));
}

void main() {
  __vybeMain();
  __check('true');
}
