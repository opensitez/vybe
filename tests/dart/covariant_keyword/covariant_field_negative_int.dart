// vybe-test: dart/covariant_keyword/covariant_field_negative_int
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

class Signed {
  num get val => 0;
}
class Negative extends Signed {
  @override
  covariant int val = -5;
}
void __vybeMain() {
  __p(Negative().val);
}

void main() {
  __vybeMain();
  __check('-5');
}
