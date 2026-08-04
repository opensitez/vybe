// vybe-test: dart/const_deep/const_class_field_initialized_at_declaration
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

class Defaults {
  static const String version = '2.0';
  static const int build = 42;
}
void __vybeMain() {
  __p(Defaults.version);
  __p(Defaults.build);
}

void main() {
  __vybeMain();
  __check('2.0\n42');
}
