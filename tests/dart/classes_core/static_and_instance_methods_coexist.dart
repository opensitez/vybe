// vybe-test: dart/classes_core/static_and_instance_methods_coexist
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

class Mix {
  int inst() {
    return 1;
  }
  static int stat() {
    return 2;
  }
}
void __vybeMain() {
  __p(Mix().inst() + Mix.stat());
}

void main() {
  __vybeMain();
  __check('3');
}
