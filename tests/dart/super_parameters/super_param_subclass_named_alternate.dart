// vybe-test: dart/super_parameters/super_param_subclass_named_alternate
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
class Rect extends Base {
  Rect(super.w, super.h);
  Rect.square(int side) : super(side, side);
}
void __vybeMain() {
  __p(Rect.square(4).w + Rect.square(4).h);
}

void main() {
  __vybeMain();
  __check('8');
}
