// vybe-test: dart/static_members/static_method_with_default_optional_param
// origin: languages/dart/tests/dart/test_static_members.rs

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

class Opt {
  static int mul(int a, [int b = 2]) {
    return a * b;
  }
}
void __vybeMain() {
  __p(Opt.mul(7));
}

void main() {
  __vybeMain();
  __check('14');
}
