// vybe-test: dart/covariant_keyword/covariant_param_list_element_type
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

class Container {
  void add(List<Object> items) {}
}
class IntContainer extends Container {
  @override
  void add(covariant List<int> items) {
    __p(items.length);
  }
}
void __vybeMain() {
  IntContainer().add([1, 2, 3]);
}

void main() {
  __vybeMain();
  __check('3');
}
