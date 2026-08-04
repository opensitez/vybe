// vybe-test: dart/class_modifiers/interface_class_method_via_interface_typed_variable
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

interface class Runner {
  int pace();
}
class Sprinter implements Runner {
  @override
  int pace() {
    return 12;
  }
}
void __vybeMain() {
  Runner r = Sprinter();
  __p(r.pace());
}

void main() {
  __vybeMain();
  __check('12');
}
