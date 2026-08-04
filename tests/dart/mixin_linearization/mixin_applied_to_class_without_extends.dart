// vybe-test: dart/mixin_linearization/mixin_applied_to_class_without_extends
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

mixin Solo {
  int one() {
    return 1;
  }
}
class Plain with Solo {}
void __vybeMain() {
  __p(Plain().one());
}

void main() {
  __vybeMain();
  __check('1');
}
