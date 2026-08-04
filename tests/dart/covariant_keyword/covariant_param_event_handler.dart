// vybe-test: dart/covariant_keyword/covariant_param_event_handler
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

class Event {}
class Click extends Event {
  int x;
  Click(this.x);
}
class Listener {
  void on(Event e) {}
}
class ClickListener extends Listener {
  @override
  void on(covariant Click e) {
    __p(e.x);
  }
}
void __vybeMain() {
  ClickListener().on(Click(99));
}

void main() {
  __vybeMain();
  __check('99');
}
