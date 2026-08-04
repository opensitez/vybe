// vybe-test: dart/hashcode_identity/identical_custom_class_same_instance
// origin: languages/dart/tests/dart/test_hashcode_identity.rs

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
  int id;
  Token(this.id);
}
void __vybeMain() {
  var t = Token(1);
  __p(identical(t, t));
}

void main() {
  __vybeMain();
  __check('true');
}
