// vybe-test: dart/null_operators/nullable_required_named_mixed_with_optional_default
// origin: languages/dart/tests/dart/test_null_operators.rs

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

void connect({required String? host, int port = 8080}) {
  __p('${host ?? 'localhost'}:$port');
}
void __vybeMain() {
  connect(host: null);
}

void main() {
  __vybeMain();
  __check('localhost:8080');
}
