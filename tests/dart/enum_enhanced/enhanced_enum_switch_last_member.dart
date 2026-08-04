// vybe-test: dart/enum_enhanced/enhanced_enum_switch_last_member
// origin: languages/dart/tests/dart/test_enum_enhanced.rs

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

enum ABC { a, b, c }
String pick(ABC v) {
  switch (v) {
    case ABC.a:
      return 'first';
    case ABC.b:
      return 'mid';
    case ABC.c:
      return 'last';
  }
}
void __vybeMain() {
  __p(pick(ABC.c));
}

void main() {
  __vybeMain();
  __check('last');
}
