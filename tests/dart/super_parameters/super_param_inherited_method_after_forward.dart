// vybe-test: dart/super_parameters/super_param_inherited_method_after_forward
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
  String tag;
  Base(this.tag);
  String label() {
    return 'base:$tag';
  }
}
class Sub extends Base {
  Sub(super.tag);
}
void __vybeMain() {
  __p(Sub('x').label());
}

void main() {
  __vybeMain();
  __check('base:x');
}
