// vybe-test: dart/interfaces_core/implements_interface_using_instance_state
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

abstract class Counter {
  int count();
}
class Tally implements Counter {
  int _n = 0;
  void inc() {
    _n = _n + 1;
  }
  int count() {
    return _n;
  }
}
void __vybeMain() {
  var t = Tally();
  t.inc();
  t.inc();
  __p(t.count());
}

void main() {
  __vybeMain();
  __check('2');
}
