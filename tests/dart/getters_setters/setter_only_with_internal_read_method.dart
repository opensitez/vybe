// vybe-test: dart/getters_setters/setter_only_with_internal_read_method
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

class WriteOnly {
  int _token = 0;
  set token(int v) {
    _token = v;
  }
  int peek() {
    return _token;
  }
}
void __vybeMain() {
  var w = WriteOnly();
  w.token = 42;
  __p(w.peek());
}

void main() {
  __vybeMain();
  __check('42');
}
