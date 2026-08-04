// vybe-test: dart/extensions/list_method_copy_reversed_order
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

extension ListFlip on List<int> {
  List<int> reversedCopy() => reversed.toList();
}
void __vybeMain() {
  __p([1, 2, 3].reversedCopy().join(','));
}

void main() {
  __vybeMain();
  __check('3,2,1');
}
