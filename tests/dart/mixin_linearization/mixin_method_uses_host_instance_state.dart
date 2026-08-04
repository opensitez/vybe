// vybe-test: dart/mixin_linearization/mixin_method_uses_host_instance_state
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

mixin Stateful {
  int ticks = 0;
  void tick() {
    ticks = ticks + 1;
  }
}
class Clock with Stateful {}
void __vybeMain() {
  var c = Clock();
  c.tick();
  c.tick();
  __p(c.ticks);
}

void main() {
  __vybeMain();
  __check('2');
}
