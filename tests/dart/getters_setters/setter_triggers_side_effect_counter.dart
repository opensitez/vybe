// vybe-test: dart/getters_setters/setter_triggers_side_effect_counter
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class Tracker {
  int _writes = 0;
  int _val = 0;
  int get writes {
    return _writes;
  }
  set val(int v) {
    _writes = _writes + 1;
    _val = v;
  }
}
void __vybeMain() {
  var t = Tracker();
  t.val = 1;
  t.val = 2;
  __p(t.writes);
}

void main() {
  __vybeMain();
  __check('2');
}
