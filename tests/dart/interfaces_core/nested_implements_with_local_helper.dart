// vybe-test: dart/interfaces_core/nested_implements_with_local_helper
// origin: languages/dart/tests/dart/test_interfaces_core.rs

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

abstract class Format {
  String fmt(int n);
}
class NumFmt implements Format {
  String fmt(int n) {
    return 'n=$n';
  }
}
void __vybeMain() {
  __p(NumFmt().fmt(42));
}

void main() {
  __vybeMain();
  __check('n=42');
}
