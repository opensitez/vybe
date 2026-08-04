// vybe-test: dart/super_parameters/super_param_deep_hierarchy_preserves_type
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Entity {
  int id;
  Entity(this.id);
}
class Model extends Entity {
  Model(super.id);
}
class User extends Model {
  User(super.id);
}
void __vybeMain() {
  __p(User(7) is Entity);
}

void main() {
  __vybeMain();
  __check('true');
}
