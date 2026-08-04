// vybe-test: dart/enums_core/enum_identity_same_reference
// origin: languages/dart/tests/dart/test_enums_core.rs

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

enum Token { alpha, beta }
void __vybeMain() {
  var a = Token.alpha;
  var b = Token.alpha;
  __p(a == b);
}

void main() {
  __vybeMain();
  __check('true');
}
