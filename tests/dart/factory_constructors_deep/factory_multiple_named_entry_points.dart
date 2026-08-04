// vybe-test: dart/factory_constructors_deep/factory_multiple_named_entry_points
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

class Color {
  int r;
  int g;
  int b;
  Color(this.r, this.g, this.b);
  factory Color.black() {
    return Color(0, 0, 0);
  }
  factory Color.white() {
    return Color(255, 255, 255);
  }
}
void __vybeMain() {
  __p(Color.black().r + Color.white().r);
}

void main() {
  __vybeMain();
  __check('255');
}
