// vybe-test: dart/factory_constructors_deep/factory_registry_returns_existing_for_same_id
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

class Entity {
  static Map<int, Entity> store = {};
  int id;
  Entity._(this.id);
  factory Entity.get(int id) {
    if (store.containsKey(id)) {
      return store[id]!;
    }
    var e = Entity._(id);
    store[id] = e;
    return e;
  }
}
void __vybeMain() {
  __p(Entity.get(1) == Entity.get(1));
}

void main() {
  __vybeMain();
  __check('true');
}
