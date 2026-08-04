// vybe-test: dart/getters_setters/private_backing_field_not_accessible_by_name
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

class Secret {
  int _hidden = 99;
  int reveal() {
    return _hidden;
  }
}
void __vybeMain() {
  __p(Secret().reveal());
}

void main() {
  __vybeMain();
  __check('99');
}
