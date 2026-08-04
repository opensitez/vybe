// vybe-test: dart/interfaces_core/abstract_getter_and_method_combo
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Profile {
  String get name;
  String label();
}
class User implements Profile {
  String get name {
    return 'u1';
  }
  String label() {
    return 'user';
  }
}
void __vybeMain() {
  var u = User();
  __p(u.name + u.label());
}

void main() {
  __vybeMain();
  __check('u1user');
}
