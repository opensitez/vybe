// vybe-test: dart/class_modifiers/base_class_getter_overridden_in_subclass
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

base class Box {
  int get size {
    return 1;
  }
}
class Crate extends Box {
  @override
  int get size {
    return 10;
  }
}
void __vybeMain() {
  __p(Crate().size);
}

void main() {
  __vybeMain();
  __check('10');
}
