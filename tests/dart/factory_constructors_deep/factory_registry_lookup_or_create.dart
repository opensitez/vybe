// vybe-test: dart/factory_constructors_deep/factory_registry_lookup_or_create
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

class Service {
  static final Map<int, Service> _registry = {};
  int id;
  Service._(this.id);
  factory Service.forId(int id) {
    return _registry.putIfAbsent(id, () => Service._(id));
  }
}
void __vybeMain() {
  __p(Service.forId(9).id);
}

void main() {
  __vybeMain();
  __check('9');
}
