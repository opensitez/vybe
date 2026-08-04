// vybe-test: dart/class_modifiers/mixin_class_field_mutation_via_host
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

mixin class Counter {
  int n = 0;
  void inc() {
    n++;
  }
}
class Box with Counter {}
void __vybeMain() {
  var b = Box();
  b.inc();
  b.inc();
  __p(b.n);
}

void main() {
  __vybeMain();
  __check('2');
}
