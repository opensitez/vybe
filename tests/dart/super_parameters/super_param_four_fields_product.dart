// vybe-test: dart/super_parameters/super_param_four_fields_product
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Base {
  int w;
  int h;
  Base(this.w, this.h);
}
class Sized extends Base {
  Sized(super.w, super.h);
  int area() {
    return w * h;
  }
}
void __vybeMain() {
  __p(Sized(3, 4).area());
}

void main() {
  __vybeMain();
  __check('12');
}
