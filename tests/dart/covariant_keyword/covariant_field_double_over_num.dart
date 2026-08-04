// vybe-test: dart/covariant_keyword/covariant_field_double_over_num
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class NumVal {
  num get reading => 0.0;
}
class DoubleVal extends NumVal {
  @override
  covariant double reading = 3.14;
}
void __vybeMain() {
  __p(DoubleVal().reading > 3.0);
}

void main() {
  __vybeMain();
  __check('true');
}
