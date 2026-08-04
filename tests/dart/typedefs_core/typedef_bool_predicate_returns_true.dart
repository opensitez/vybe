// vybe-test: dart/typedefs_core/typedef_bool_predicate_returns_true
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef Filter = bool Function(int);
bool isPositive(int n) {
  return n > 0;
}
void __vybeMain() {
  Filter pred = isPositive;
  __p(pred(3));
}

void main() {
  __vybeMain();
  __check('true');
}
