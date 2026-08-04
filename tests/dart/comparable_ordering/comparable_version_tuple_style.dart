// vybe-test: dart/comparable_ordering/comparable_version_tuple_style
// origin: languages/dart/tests/dart/test_comparable_ordering.rs

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

class Ver implements Comparable<Ver> {
  int major;
  int minor;
  Ver(this.major, this.minor);
  int compareTo(Ver other) {
    var c = major.compareTo(other.major);
    return c != 0 ? c : minor.compareTo(other.minor);
  }
}
void __vybeMain() {
  __p(Ver(2, 0).compareTo(Ver(1, 9)));
}

void main() {
  __vybeMain();
  __check('1');
}
