// vybe-test: dart/covariant_keyword/covariant_param_account_hierarchy
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

class Account {
  int balance;
  Account(this.balance);
}
class Savings extends Account {
  Savings(int b) : super(b);
}
class Bank {
  void deposit(Account a, int amt) {}
}
class SavingsBank extends Bank {
  @override
  void deposit(covariant Savings s, int amt) {
    __p(s.balance + amt);
  }
}
void __vybeMain() {
  SavingsBank().deposit(Savings(100), 50);
}

void main() {
  __vybeMain();
  __check('150');
}
