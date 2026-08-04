// vybe-test: dart/covariant_keyword/covariant_param_bool_subclass
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

class Flag {
  bool on;
  Flag(this.on);
}
class Switch extends Flag {
  Switch(bool v) : super(v);
}
class Panel {
  void flip(Flag f) {}
}
class SwitchPanel extends Panel {
  @override
  void flip(covariant Switch s) {
    __p(s.on);
  }
}
void __vybeMain() {
  SwitchPanel().flip(Switch(true));
}

void main() {
  __vybeMain();
  __check('true');
}
