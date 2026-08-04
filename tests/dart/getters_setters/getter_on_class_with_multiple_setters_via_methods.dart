// vybe-test: dart/getters_setters/getter_on_class_with_multiple_setters_via_methods
// origin: languages/dart/tests/dart/test_getters_setters.rs

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

class Account {
  int _balance = 100;
  int get balance {
    return _balance;
  }
  void deposit(int amount) {
    _balance = _balance + amount;
  }
  void withdraw(int amount) {
    _balance = _balance - amount;
  }
}
void __vybeMain() {
  var a = Account();
  a.deposit(50);
  a.withdraw(30);
  __p(a.balance);
}

void main() {
  __vybeMain();
  __check('120');
}
