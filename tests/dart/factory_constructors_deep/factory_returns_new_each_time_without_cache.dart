// vybe-test: dart/factory_constructors_deep/factory_returns_new_each_time_without_cache
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

class Box {
  int id;
  static int _next = 0;
  Box._(this.id);
  factory Box() {
    _next = _next + 1;
    return Box._(_next);
  }
}
void __vybeMain() {
  __p(Box().id + Box().id);
}

void main() {
  __vybeMain();
  __check('3');
}
