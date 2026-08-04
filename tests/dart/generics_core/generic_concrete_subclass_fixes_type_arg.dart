// vybe-test: dart/generics_core/generic_concrete_subclass_fixes_type_arg
// origin: languages/dart/tests/dart/test_generics_core.rs

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

class Base<T> {
  T value;
  Base(this.value);
}
class IntBox extends Base<int> {
  IntBox(int v) : super(v);
}
void __vybeMain() {
  var box = IntBox(15);
  __p(box.value);
}

void main() {
  __vybeMain();
  __check('15');
}
