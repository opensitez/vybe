// vybe-test: dart/extensions/iterable_method_first_matching_or_zero
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

extension IterFind on Iterable<int> {
  int firstGreaterThan(int threshold) {
    for (var n in this) {
      if (n > threshold) return n;
    }
    return 0;
  }
}
void __vybeMain() {
  __p([1, 5, 3].firstGreaterThan(2));
}

void main() {
  __vybeMain();
  __check('5');
}
