// vybe-test: dart/extensions/int_getter_is_odd_on_seven
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

extension IntParity on int {
  bool get isOdd => this % 2 != 0;
}
void __vybeMain() {
  __p(7.isOdd);
}

void main() {
  __vybeMain();
  __check('true');
}
