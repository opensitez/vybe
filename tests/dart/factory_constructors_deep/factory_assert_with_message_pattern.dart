// vybe-test: dart/factory_constructors_deep/factory_assert_with_message_pattern
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

class Port {
  int number;
  Port._(this.number);
  factory Port(int n) {
    assert(n > 0, 'port must be positive');
    return Port._(n);
  }
}
void __vybeMain() {
  __p(Port(443).number);
}

void main() {
  __vybeMain();
  __check('443');
}
