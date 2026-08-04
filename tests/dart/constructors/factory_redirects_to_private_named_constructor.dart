// vybe-test: dart/constructors/factory_redirects_to_private_named_constructor
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

class Token {
  String text;
  Token._(this.text);
  factory Token.fromText(String t) {
    return Token._(t);
  }
}
void __vybeMain() {
  __p(Token.fromText('ok').text);
}

void main() {
  __vybeMain();
  __check('ok');
}
