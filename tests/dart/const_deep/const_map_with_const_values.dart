// vybe-test: dart/const_deep/const_map_with_const_values
// origin: languages/dart/tests/dart/test_const_deep.rs

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

class Code {
  final int n;
  const Code(this.n);
}
void __vybeMain() {
  const codes = {'a': Code(1), 'b': Code(2)};
  __p(codes['b']!.n);
}

void main() {
  __vybeMain();
  __check('2');
}
