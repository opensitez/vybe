// vybe-test: dart/cascades/cascade_custom_setter_chain_updates_backing_field
// origin: languages/dart/tests/dart/test_cascades.rs

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
  int _size = 0;
  set size(int v) { _size = v; }
  int get size => _size;
}
void __vybeMain() {
  var crate = Box();
  crate..size = 10..size = 25;
  __p(crate.size);
}

void main() {
  __vybeMain();
  __check('25');
}
