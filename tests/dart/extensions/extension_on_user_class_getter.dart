// vybe-test: dart/extensions/extension_on_user_class_getter
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

class Box { int v = 4; }
extension BoxX on Box { int get tripled => v * 3; }
void __vybeMain() {
  __p(Box().tripled);
}

void main() {
  __vybeMain();
  __check('12');
}
