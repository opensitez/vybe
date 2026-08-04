// vybe-test: dart/generic_inference/infer_var_type_from_generic_method_return
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

class Cell<T> {
  T value;
  Cell(this.value);
  Cell<R> map<R>(R Function(T) fn) {
    return Cell(fn(value));
  }
}
void __vybeMain() {
  var c = Cell(5);
  var mapped = c.map((n) => n.toString());
  __p(mapped.value);
}

void main() {
  __vybeMain();
  __check('5');
}
