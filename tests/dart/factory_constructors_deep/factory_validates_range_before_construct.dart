// vybe-test: dart/factory_constructors_deep/factory_validates_range_before_construct
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

class Percent {
  int value;
  Percent._(this.value);
  factory Percent(int v) {
    assert(v >= 0 && v <= 100);
    return Percent._(v);
  }
}
void __vybeMain() {
  __p(Percent(75).value);
}

void main() {
  __vybeMain();
  __check('75');
}
