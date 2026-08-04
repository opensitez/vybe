// vybe-test: dart/operator_overloading/operator_equals_single_field
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

class Id {
  int value;
  Id(this.value);
  bool operator ==(Object other) {
    return other is Id && value == other.value;
  }
}
void __vybeMain() {
  __p(Id(7) == Id(7));
}

void main() {
  __vybeMain();
  __check('true');
}
