// vybe-test: dart/factory_constructors_deep/factory_assert_validates_positive_input
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

class Age {
  int years;
  Age._(this.years);
  factory Age(int y) {
    assert(y >= 0);
    return Age._(y);
  }
}
void __vybeMain() {
  __p(Age(30).years);
}

void main() {
  __vybeMain();
  __check('30');
}
