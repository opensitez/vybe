// vybe-test: dart/operator_overloading/operator_equals_int_wrapped_value
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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
  int inner;
  Wrap(this.inner);
  bool operator ==(Object other) {
    return other is Wrap && inner == other.inner;
  }
}
void __vybeMain() {
  __p(Wrap(0) == Wrap(0));
}

void main() {
  __vybeMain();
  __check('true');
}
