// vybe-test: dart/generic_inference/infer_generic_method_keep_where_on_ints
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

class Holder<T> {
  List<T> items;
  Holder(this.items);
  List<T> keepWhere(bool Function(T) test) {
    return items.where(test).toList();
  }
}
void __vybeMain() {
  var h = Holder([1, 2, 3, 4]);
  var evens = h.keepWhere((n) => n % 2 == 0);
  __p(evens.join(','));
}

void main() {
  __vybeMain();
  __check('2,4');
}
