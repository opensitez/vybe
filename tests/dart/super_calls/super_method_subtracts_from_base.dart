// vybe-test: dart/super_calls/super_method_subtracts_from_base
// origin: languages/dart/tests/dart/test_super_calls.rs

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

class Wallet {
  int balance() {
    return 20;
  }
}
class Fee extends Wallet {
  int balance() {
    return super.balance() - 5;
  }
}
void __vybeMain() {
  __p(Fee().balance());
}

void main() {
  __vybeMain();
  __check('15');
}
