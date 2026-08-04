// vybe-test: dart/super_calls/super_method_used_in_arithmetic_expression
// origin: languages/dart/tests/dart/test_super_calls.rs

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

class Meter {
  int read() {
    return 6;
  }
}
class Adjusted extends Meter {
  int read() {
    return super.read() * 2 + 1;
  }
}
void __vybeMain() {
  __p(Adjusted().read());
}

void main() {
  __vybeMain();
  __check('13');
}
