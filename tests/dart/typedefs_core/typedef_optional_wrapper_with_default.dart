// vybe-test: dart/typedefs_core/typedef_optional_wrapper_with_default
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

typedef MaybeFn = int Function(int);
int runOrZero(MaybeFn? fn, int input) {
  if (fn == null) {
    return 0;
  }
  return fn(input);
}
int doubleIt(int n) {
  return n * 2;
}
void __vybeMain() {
  __p(runOrZero(doubleIt, 4));
  __p(runOrZero(null, 4));
}

void main() {
  __vybeMain();
  __check('8\n0');
}
