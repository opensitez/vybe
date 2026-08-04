// vybe-test: dart/generics_core/generic_nullable_field_on_cache
// origin: languages/dart/tests/dart/test_generics_core.rs

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

class Cache<T> {
  T? _value;
  T? get value {
    return _value;
  }
  void store(T v) {
    _value = v;
  }
}
void __vybeMain() {
  var c = Cache<int>();
  __p(c.value);
  c.store(5);
  __p(c.value);
}

void main() {
  __vybeMain();
  __check('null\n5');
}
