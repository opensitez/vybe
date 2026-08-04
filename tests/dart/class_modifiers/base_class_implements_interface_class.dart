// vybe-test: dart/class_modifiers/base_class_implements_interface_class
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

interface class Describable {
  String describe();
}
base class Entity implements Describable {
  @override
  String describe() {
    return 'entity';
  }
}
class User extends Entity {
  @override
  String describe() {
    return 'user';
  }
}
void __vybeMain() {
  __p(User().describe());
}

void main() {
  __vybeMain();
  __check('user');
}
