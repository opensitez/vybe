// vybe-test: dart/covariant_keyword/covariant_field_int_over_num_getter
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

class NumBox {
  num get value => 0;
}
class IntBox extends NumBox {
  @override
  covariant int value = 7;
}
void __vybeMain() {
  __p(IntBox().value);
}

void main() {
  __vybeMain();
  __check('7');
}
