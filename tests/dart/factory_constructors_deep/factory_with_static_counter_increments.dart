// vybe-test: dart/factory_constructors_deep/factory_with_static_counter_increments
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

class Ticket {
  static int serial = 0;
  int number;
  Ticket(this.number);
  factory Ticket.next() {
    serial = serial + 1;
    return Ticket(serial);
  }
}
void __vybeMain() {
  __p(Ticket.next().number);
  __p(Ticket.next().number);
}

void main() {
  __vybeMain();
  __check('1\n2');
}
