// vybe-test: dart/named_parameters/constructor_named_params_with_defaults
// origin: languages/dart/tests/dart/test_named_parameters.rs

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
  final String name;
  final int age;
  User({this.name = 'anon', this.age = 0});
}
void __vybeMain() {
  var u = User();
  __p('${u.name}:${u.age}');
}

void main() {
  __vybeMain();
  __check('anon:0');
}
