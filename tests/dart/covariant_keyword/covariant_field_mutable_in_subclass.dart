// vybe-test: dart/covariant_keyword/covariant_field_mutable_in_subclass
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

class Base {
  num get n => 1;
}
class Sub extends Base {
  @override
  covariant int n = 5;
}
void __vybeMain() {
  var s = Sub();
  s.n = 10;
  __p(s.n);
}

void main() {
  __vybeMain();
  __check('10');
}
