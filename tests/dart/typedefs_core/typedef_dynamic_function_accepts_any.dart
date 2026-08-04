// vybe-test: dart/typedefs_core/typedef_dynamic_function_accepts_any
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

typedef DynFn = dynamic Function(dynamic);
dynamic echo(dynamic value) {
  return value;
}
void __vybeMain() {
  DynFn fn = echo;
  __p(fn('x'));
  __p(fn(9));
}

void main() {
  __vybeMain();
  __check('x\n9');
}
