// vybe-test: dart/covariant_keyword/covariant_param_method_count_items
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Collection {
  int count(List<Object> xs) {
    return xs.length;
  }
}
class IntCollection extends Collection {
  @override
  int count(covariant List<int> xs) {
    return xs.length + 1;
  }
}
void __vybeMain() {
  __p(IntCollection().count([1, 2]));
}

void main() {
  __vybeMain();
  __check('3');
}
