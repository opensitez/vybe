// vybe-test: dart/generic_inference/infer_set_element_type_from_literals
// origin: languages/dart/tests/dart/test_generic_inference.rs

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

void __vybeMain() {
  var tags = {'dart', 'flutter', 'dart'};
  __p(tags.length);
  __p(tags.contains('dart'));
}

void main() {
  __vybeMain();
  __check('2\ntrue');
}
