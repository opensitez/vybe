// vybe-test: dart/extensions/iterable_getter_has_any_item
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension IterAny on Iterable<int> {
  bool get hasItems => !isEmpty;
}
void __vybeMain() {
  __p([1].hasItems);
}

void main() {
  __vybeMain();
  __check('true');
}
