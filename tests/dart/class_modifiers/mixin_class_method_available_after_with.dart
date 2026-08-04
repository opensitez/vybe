// vybe-test: dart/class_modifiers/mixin_class_method_available_after_with
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

mixin class Loggable {
  String tag() {
    return 'log';
  }
}
class App with Loggable {}
void __vybeMain() {
  __p(App().tag());
}

void main() {
  __vybeMain();
  __check('log');
}
