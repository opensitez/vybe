// vybe-test: dart/class_modifiers/sealed_subtype_list_in_same_library
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Token {}
class Alpha extends Token {}
class Beta extends Token {}
class Gamma extends Token {}
int rank(Token t) {
  switch (t) {
    case Alpha():
      return 1;
    case Beta():
      return 2;
    case Gamma():
      return 3;
  }
}
void __vybeMain() {
  __p(rank(Beta()));
}

void main() {
  __vybeMain();
  __check('2');
}
