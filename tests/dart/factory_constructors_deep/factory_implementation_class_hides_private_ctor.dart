// vybe-test: dart/factory_constructors_deep/factory_implementation_class_hides_private_ctor
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Hidden {
  int code;
  Hidden._(this.code);
  factory Hidden.open(int c) {
    return Hidden._(c);
  }
}
void __vybeMain() {
  __p(Hidden.open(42).code);
}

void main() {
  __vybeMain();
  __check('42');
}
