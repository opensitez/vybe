// vybe-test: dart/abstract_members/abstract_getter_computed_from_field
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Sized {
  int get width;
  int get height;
  int area() {
    return width * height;
  }
}
class Rect extends Sized {
  int w;
  int h;
  Rect(this.w, this.h);
  int get width => w;
  int get height => h;
}
void __vybeMain() {
  __p(Rect(3, 4).area());
}

void main() {
  __vybeMain();
  __check('12');
}
