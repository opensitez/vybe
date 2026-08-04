// vybe-test: dart/const_deep/static_const_field_access
// origin: languages/dart/tests/dart/test_const_deep.rs

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
  static const int timeout = 30;
  static const String host = 'localhost';
}
void __vybeMain() {
  __p(Config.timeout);
  __p(Config.host);
}

void main() {
  __vybeMain();
  __check('30\nlocalhost');
}
