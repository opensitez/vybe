// vybe-test: dart/constructors/const_constructor_field_access
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

class ConstBox {
  final int w;
  final int h;
  const ConstBox(this.w, this.h);
  int get area {
    return w * h;
  }
}
void __vybeMain() {
  const b = ConstBox(2, 5);
  __p(b.area);
}

void main() {
  __vybeMain();
  __check('10');
}
