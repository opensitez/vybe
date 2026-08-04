// vybe-test: dart/factory_constructors_deep/factory_from_json_reads_string_field
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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
  int id;
  String name;
  User._(this.id, this.name);
  factory User.fromJson(Map<String, dynamic> json) {
    return User._(json['id'], json['name']);
  }
}
void __vybeMain() {
  var u = User.fromJson({'id': 1, 'name': 'Bob'});
  __p(u.name);
}

void main() {
  __vybeMain();
  __check('Bob');
}
