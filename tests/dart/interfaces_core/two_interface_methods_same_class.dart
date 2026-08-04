// vybe-test: dart/interfaces_core/two_interface_methods_same_class
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

abstract class Reader {
  int read();
}
abstract class Writer {
  int write();
}
class RW implements Reader, Writer {
  int read() {
    return 10;
  }
  int write() {
    return 20;
  }
}
void __vybeMain() {
  var rw = RW();
  __p(rw.read() + rw.write());
}

void main() {
  __vybeMain();
  __check('30');
}
