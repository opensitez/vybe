// vybe-test: dart/enum_enhanced/enhanced_enum_const_with_body_logic
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

enum Version {
  v1(1),
  v2(2);
  final int major;
  const Version(this.major);
  int next() {
    return major + 1;
  }
}
void __vybeMain() {
  __p(Version.v1.next());
}

void main() {
  __vybeMain();
  __check('2');
}
