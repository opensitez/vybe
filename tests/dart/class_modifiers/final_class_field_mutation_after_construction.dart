// vybe-test: dart/class_modifiers/final_class_field_mutation_after_construction
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

final class Acc {
  int total = 0;
  void add(int n) {
    total = total + n;
  }
}
void __vybeMain() {
  var a = Acc();
  a.add(3);
  a.add(4);
  __p(a.total);
}

void main() {
  __vybeMain();
  __check('7');
}
