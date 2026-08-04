// vybe-test: dart/interfaces_core/interface_void_method_implementation
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

abstract class Action {
  void run();
}
class Job implements Action {
  int done = 0;
  void run() {
    done = 1;
  }
}
void __vybeMain() {
  var j = Job();
  j.run();
  __p(j.done);
}

void main() {
  __vybeMain();
  __check('1');
}
