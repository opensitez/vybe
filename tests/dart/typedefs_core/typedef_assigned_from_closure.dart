// vybe-test: dart/typedefs_core/typedef_assigned_from_closure
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

typedef Doubler = int Function(int);
void __vybeMain() {
  Doubler fn = (int n) => n * 2;
  __p(fn(6));
}

void main() {
  __vybeMain();
  __check('12');
}
