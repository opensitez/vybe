// vybe-test: dart/generics_core/generic_class_type_inference_on_constructor
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

class Box<T> {
  T value;
  Box(this.value);
}
void __vybeMain() {
  var b = Box(42);
  __p(b.value);
}

void main() {
  __vybeMain();
  __check('42');
}
