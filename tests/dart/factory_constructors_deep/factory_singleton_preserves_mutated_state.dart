// vybe-test: dart/factory_constructors_deep/factory_singleton_preserves_mutated_state
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

class Cache {
  static Cache? _one;
  int count = 0;
  Cache._();
  factory Cache() {
    _one ??= Cache._();
    return _one!;
  }
}
void __vybeMain() {
  var a = Cache();
  a.count = 5;
  __p(Cache().count);
}

void main() {
  __vybeMain();
  __check('5');
}
