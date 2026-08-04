// vybe-test: dart/constructors/factory_constructor_can_return_cached_instance
// origin: languages/dart/tests/dart/test_constructors.rs

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
  int id;
  Cache._(this.id);
  factory Cache() {
    _one ??= Cache._(1);
    return _one!;
  }
}
void __vybeMain() {
  var a = Cache();
  var b = Cache();
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true');
}
