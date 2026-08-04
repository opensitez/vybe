// vybe-test: dart/static_members/static_getter_bool_flag
// origin: languages/dart/tests/dart/test_static_members.rs

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

class Switch {
  static bool _on = false;
  static bool get on => _on;
}
void __vybeMain() {
  __p(Switch.on);
}

void main() {
  __vybeMain();
  __check('false');
}
