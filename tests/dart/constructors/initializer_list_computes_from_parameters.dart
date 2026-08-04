// vybe-test: dart/constructors/initializer_list_computes_from_parameters
// origin: languages/dart/tests/dart/test_constructors.rs

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

class Span {
  int start;
  int end;
  Span(int len) : start = 0, end = len;
}
void __vybeMain() {
  __p(Span(5).end);
}

void main() {
  __vybeMain();
  __check('5');
}
