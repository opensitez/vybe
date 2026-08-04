// vybe-test: dart/factory_constructors_deep/factory_redirect_arrow_to_private_constructor
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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
  final String value;
  Token._(this.value);
  factory Token(String v) => Token._(v);
}
void __vybeMain() {
  __p(Token('secret').value);
}

void main() {
  __vybeMain();
  __check('secret');
}
