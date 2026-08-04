// vybe-test: dart/classes_core/method_uses_this_to_disambiguate_field
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

class Wrap {
  int value = 3;
  int bump() {
    return this.value + 1;
  }
}
void __vybeMain() {
  __p(Wrap().bump());
}

void main() {
  __vybeMain();
  __check('4');
}
