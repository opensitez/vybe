// vybe-test: dart/classes_core/instance_method_reads_static_field
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

class Reader {
  static int shared = 6;
  int readShared() {
    return shared;
  }
}
void __vybeMain() {
  __p(Reader().readShared());
}

void main() {
  __vybeMain();
  __check('6');
}
