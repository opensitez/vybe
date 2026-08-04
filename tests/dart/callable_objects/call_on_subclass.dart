// vybe-test: dart/callable_objects/call_on_subclass
// origin: languages/dart/tests/dart/test_callable_objects.rs

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
  int call(int n) {
    return n;
  }
}
class DoubleCall extends Base {
  @override
  int call(int n) {
    return n * 2;
  }
}
void __vybeMain() {
  __p(DoubleCall()(7));
}

void main() {
  __vybeMain();
  __check('14');
}
