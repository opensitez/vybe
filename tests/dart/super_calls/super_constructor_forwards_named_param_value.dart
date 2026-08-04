// vybe-test: dart/super_calls/super_constructor_forwards_named_param_value
// origin: languages/dart/tests/dart/test_super_calls.rs

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

class User {
  String role;
  User(this.role);
}
class Admin extends User {
  Admin() : super('admin');
}
void __vybeMain() {
  __p(Admin().role);
}

void main() {
  __vybeMain();
  __check('admin');
}
