// vybe-test: dart/field_initializers/final_field_set_only_in_initializer_list
// origin: languages/dart/tests/dart/test_field_initializers.rs

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
  final String code;
  Token(String c) : code = c;
}
void __vybeMain() {
  __p(Token('abc').code);
}

void main() {
  __vybeMain();
  __check('abc');
}
