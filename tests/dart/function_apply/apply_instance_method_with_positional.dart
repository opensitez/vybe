// vybe-test: dart/function_apply/apply_instance_method_with_positional
// origin: languages/dart/tests/dart/test_function_apply.rs

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

class Calc {
  int add(int a, int b) {
    return a + b;
  }
}
void __vybeMain() {
  var c = Calc();
  __p(Function.apply(c.add, [2, 3]));
}

void main() {
  __vybeMain();
  __check('5');
}
