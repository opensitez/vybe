// vybe-test: dart/classes_core/static_method_calls_other_static_method
// origin: languages/dart/tests/dart/test_classes_core.rs

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

class Chain {
  static int step1(int n) {
    return n + 1;
  }
  static int step2(int n) {
    return step1(n) * 2;
  }
}
void __vybeMain() {
  __p(Chain.step2(4));
}

void main() {
  __vybeMain();
  __check('10');
}
