// vybe-test: dart/callable_objects/call_with_mixed_positional_and_named
// origin: languages/dart/tests/dart/test_callable_objects.rs

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

class Mixer {
  int call(int base, {int bonus = 0}) {
    return base + bonus;
  }
}
void __vybeMain() {
  __p(Mixer()(5, bonus: 2));
}

void main() {
  __vybeMain();
  __check('7');
}
