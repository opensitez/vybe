// vybe-test: dart/class_modifiers/interface_class_multiple_implementers
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

interface class Identifiable {
  String id();
}
class User implements Identifiable {
  @override
  String id() {
    return 'u1';
  }
}
class Guest implements Identifiable {
  @override
  String id() {
    return 'g1';
  }
}
void __vybeMain() {
  __p(User().id() + Guest().id());
}

void main() {
  __vybeMain();
  __check('u1g1');
}
