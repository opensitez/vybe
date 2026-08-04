// vybe-test: dart/interfaces_core/interface_arrow_method_body
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Double {
  int twice(int n);
}
class Fast implements Double {
  int twice(int n) => n * 2;
}
void __vybeMain() {
  __p(Fast().twice(7));
}

void main() {
  __vybeMain();
  __check('14');
}
