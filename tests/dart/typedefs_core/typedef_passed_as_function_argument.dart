// vybe-test: dart/typedefs_core/typedef_passed_as_function_argument
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

typedef Mapper = int Function(int);
int applyTwice(int value, Mapper fn) {
  return fn(fn(value));
}
int inc(int n) {
  return n + 1;
}
void __vybeMain() {
  __p(applyTwice(3, inc));
}

void main() {
  __vybeMain();
  __check('5');
}
