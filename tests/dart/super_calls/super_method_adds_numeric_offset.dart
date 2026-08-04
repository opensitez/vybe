// vybe-test: dart/super_calls/super_method_adds_numeric_offset
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

class Counter {
  int value() {
    return 10;
  }
}
class Boost extends Counter {
  int value() {
    return super.value() + 5;
  }
}
void __vybeMain() {
  __p(Boost().value());
}

void main() {
  __vybeMain();
  __check('15');
}
