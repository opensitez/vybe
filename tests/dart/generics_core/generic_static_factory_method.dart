// vybe-test: dart/generics_core/generic_static_factory_method
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

class Wrapper<T> {
  T data;
  Wrapper(this.data);
  static Wrapper<T> of<T>(T data) {
    return Wrapper(data);
  }
}
void __vybeMain() {
  var w = Wrapper.of(42);
  __p(w.data);
}

void main() {
  __vybeMain();
  __check('42');
}
