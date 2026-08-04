// vybe-test: dart/getters_setters/getter_with_nullable_backing_coalesces
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class MaybeName {
  String? _name;
  String get display {
    return _name ?? 'anonymous';
  }
  set display(String v) {
    _name = v;
  }
}
void __vybeMain() {
  var m = MaybeName();
  __p(m.display);
  m.display = 'Zara';
  __p(m.display);
}

void main() {
  __vybeMain();
  __check('anonymous\nZara');
}
