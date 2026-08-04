// vybe-test: dart/class_modifiers/sealed_class_switch_with_guard_like_pattern
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Num {}
class Positive extends Num {
  int v;
  Positive(this.v);
}
class Zero extends Num {}
String sign(Num n) {
  switch (n) {
    case Positive(v: var x) when x > 0:
      return 'pos';
    case Zero():
      return 'zero';
    case Positive():
      return 'other';
  }
}
void __vybeMain() {
  __p(sign(Positive(5)));
}

void main() {
  __vybeMain();
  __check('pos');
}
