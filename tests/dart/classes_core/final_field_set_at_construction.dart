// vybe-test: dart/classes_core/final_field_set_at_construction
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

class Token {
  final String value;
  Token(this.value);
}
void __vybeMain() {
  var t = Token('abc');
  __p(t.value);
}

void main() {
  __vybeMain();
  __check('abc');
}
