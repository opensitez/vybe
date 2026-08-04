// vybe-test: dart/named_parameters/forwarding_named_to_inner_function
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

void inner({required String key, int n = 0}) {
  __p('$key=$n');
}
void outer({required String key, int n = 0}) {
  inner(key: key, n: n);
}
void __vybeMain() {
  outer(key: 'id', n: 7);
}

void main() {
  __vybeMain();
  __check('id=7');
}
