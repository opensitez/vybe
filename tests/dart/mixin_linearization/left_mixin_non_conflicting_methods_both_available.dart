// vybe-test: dart/mixin_linearization/left_mixin_non_conflicting_methods_both_available
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

mixin Left {
  int leftVal() {
    return 1;
  }
}
mixin Right {
  int rightVal() {
    return 2;
  }
}
class Both with Left, Right {}
void __vybeMain() {
  __p(Both().leftVal() + Both().rightVal());
}

void main() {
  __vybeMain();
  __check('3');
}
