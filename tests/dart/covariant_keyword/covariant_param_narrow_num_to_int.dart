// vybe-test: dart/covariant_keyword/covariant_param_narrow_num_to_int
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class NumHolder {
  void store(num n) {}
}
class IntHolder extends NumHolder {
  @override
  void store(covariant int n) {}
}
void __vybeMain() {
  IntHolder().store(42);
  __p(42);
}

void main() {
  __vybeMain();
  __check('42');
}
