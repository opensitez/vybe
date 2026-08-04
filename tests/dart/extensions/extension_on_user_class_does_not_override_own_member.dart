// vybe-test: dart/extensions/extension_on_user_class_does_not_override_own_member
// origin: languages/dart/tests/dart/test_extensions.rs

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

class Box {
  int v = 5;
  int twice() => 999;
}
extension BoxX on Box { int twice() => v * 2; }
void __vybeMain() {
  __p(Box().twice());
}

void main() {
  __vybeMain();
  __check('999');
}
