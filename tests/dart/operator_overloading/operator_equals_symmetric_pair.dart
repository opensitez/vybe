// vybe-test: dart/operator_overloading/operator_equals_symmetric_pair
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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
  String key;
  Token(this.key);
  bool operator ==(Object other) {
    return other is Token && key == other.key;
  }
}
void __vybeMain() {
  var a = Token('k');
  var b = Token('k');
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true');
}
