// vybe-test: dart/interfaces_core/interface_with_bool_return
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

abstract class Check {
  bool ok();
}
class Pass implements Check {
  bool ok() {
    return true;
  }
}
void __vybeMain() {
  __p(Pass().ok());
}

void main() {
  __vybeMain();
  __check('true');
}
