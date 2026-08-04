// vybe-test: dart/abstract_members/abstract_method_called_from_concrete_method
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Tax {
  double rate();
  double apply(double amount) {
    return amount * rate();
  }
}
class SalesTax extends Tax {
  double rate() {
    return 0.1;
  }
}
void __vybeMain() {
  __p(SalesTax().apply(100.0));
}

void main() {
  __vybeMain();
  __check('10.0');
}
