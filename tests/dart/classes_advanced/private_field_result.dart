// vybe-test: dart/classes_advanced/private_field_result
// origin: languages/dart/tests/dart/test_classes_advanced.rs

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

class Account { double _balance = 0; void deposit(double v) { _balance += v; } double get balance => _balance; }
void __vybeMain() { var a = Account(); a.deposit(100); __p(a.balance); }

void main() {
  __vybeMain();
  // Damaged test repaired: `_balance` is a `double`, and dart 3.10.4 renders
  // a double as "100.0" (measured) — the expectation asserted the int
  // spelling and failed under real dart.
  __check('100.0');
}
