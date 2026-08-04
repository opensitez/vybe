// vybe-test: dart/constructors/private_generative_with_public_factory
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Secret {
  int code;
  Secret._(this.code);
  factory Secret.open(int c) {
    return Secret._(c);
  }
}
void __vybeMain() {
  __p(Secret.open(99).code);
}

void main() {
  __vybeMain();
  __check('99');
}
