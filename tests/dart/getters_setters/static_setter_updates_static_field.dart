// vybe-test: dart/getters_setters/static_setter_updates_static_field
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

class Config {
  static int _port = 8080;
  static int get port {
    return _port;
  }
  static set port(int v) {
    _port = v;
  }
}
void __vybeMain() {
  Config.port = 3000;
  __p(Config.port);
}

void main() {
  __vybeMain();
  __check('3000');
}
