// vybe-test: dart/factory_constructors_deep/factory_named_from_list_first_element
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Head {
  int value;
  Head(this.value);
  factory Head.fromList(List<int> items) {
    return Head(items.first);
  }
}
void __vybeMain() {
  __p(Head.fromList([9, 8, 7]).value);
}

void main() {
  __vybeMain();
  __check('9');
}
