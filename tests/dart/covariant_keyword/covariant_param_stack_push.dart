// vybe-test: dart/covariant_keyword/covariant_param_stack_push
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

class Stack {
  void push(Object item) {}
}
class IntStack extends Stack {
  int total = 0;
  @override
  void push(covariant int item) {
    total = total + item;
  }
}
void __vybeMain() {
  var s = IntStack();
  s.push(3);
  s.push(4);
  __p(s.total);
}

void main() {
  __vybeMain();
  __check('7');
}
