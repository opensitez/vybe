// vybe-test: dart/generics_core/generic_optional_default
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

class Cell<T> {
  T? _val;
  T readOr(T fallback) {
    return _val ?? fallback;
  }
  void write(T v) {
    _val = v;
  }
}
void __vybeMain() {
  var c = Cell<int>();
  __p(c.readOr(-1));
  c.write(8);
  __p(c.readOr(-1));
}

void main() {
  __vybeMain();
  __check('-1\n8');
}
