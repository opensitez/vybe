// vybe-test: dart/classes_core/static_field_read_via_class_name
// origin: languages/dart/tests/dart/test_classes_core.rs

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
  static int version = 2;
}
void __vybeMain() {
  __p(Config.version);
}

void main() {
  __vybeMain();
  __check('2');
}
