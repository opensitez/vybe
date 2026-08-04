// vybe-test: dart/comparable_ordering/comparable_version_sort_order
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
  var vers = [Ver(2, 1), Ver(1, 10), Ver(2, 0)];
  vers.sort();
  __p(vers[0].major);
  __p(vers[0].minor);
  __p(vers[2].major);
}

void main() {
  __vybeMain();
  __check('1\n10\n2');
}
