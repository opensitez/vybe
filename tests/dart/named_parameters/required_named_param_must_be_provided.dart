// vybe-test: dart/named_parameters/required_named_param_must_be_provided
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

void connect({required String host}) {
  __p(host);
}
void __vybeMain() {
  connect(host: 'localhost');
}

void main() {
  __vybeMain();
  __check('localhost');
}
